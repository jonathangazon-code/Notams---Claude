using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// One flight-schedule row, trimmed to what the conflict scan needs.
		private struct FsFlight
		{
			public string Callsign, Reg, Origin, Dest;
			public DateTime Std, Sta;
			public bool HasStd, HasSta;
		}

		void ConflictTabEnter(object sender, EventArgs e) { Build_Conflict_Report(); }

		// Cross-references every Kept, impact-classified NOTAM against FlightSchedule:
		// a conflict exists when the NOTAM's validity window overlaps the ±window (Admin
		// tab, _conflictWindowHours) around a flight's STD (station used as Origin) or STA
		// (station used as Dest). One section per impact code (A/N/C/D/F, same severity
		// order as SuggestedSingleCode), each listing the airport card + matching flights +
		// full NOTAM text for every conflict found; SUP is out of scope for this tab.
		void Build_Conflict_Report()
		{
			Dictionary<string, string> unused;
			Web_Conflict.DocumentText = BuildConflictReportHtml(false, false, out unused);
		}

		// Builds the full Conflict-report HTML document — used both to populate the Conflict
		// tab's WebBrowser and as the Send Reports email body (MainForm.Email.cs), so the two
		// stay identical rather than drifting into two separately-maintained renderings.
		//
		// rasterDiagrams/cidImages control how the runway diagrams (normally VML, only
		// understood by this app's own IE7-mode WebBrowser control) are embedded:
		//   - rasterDiagrams=false: VML, as before (live Conflict tab).
		//   - rasterDiagrams=true, cidImages=false: PNG as a data-URI <img> (safe for
		//     wkhtmltopdf/WebKit, not used by this method's own callers today but mirrors
		//     the NOTAM Report tab's usage of BuildAirportHeaderHtml).
		//   - rasterDiagrams=true, cidImages=true: PNG written to a temp file and embedded as
		//     <img src="cid:...">, the only image form Outlook's Word engine actually renders
		//     in an HTML email body — inlineImages (cid -> temp file path) is what
		//     MainForm.Email.cs attaches to the mail item after building the body.
		public string BuildConflictReportHtml(bool rasterDiagrams, bool cidImages, out Dictionary<string, string> inlineImages)
		{
			inlineImages = new Dictionary<string, string>();

			EnsureArchiveConfig();
			List<FsFlight> flights = LoadFsFlights();

			// The whole tab is capped to the next 7 days — a NOTAM/flight pairing further
			// out than that isn't shown, regardless of how far ahead FlightSchedule itself
			// happens to reach (it can extend well past 7 days via the CSV fallback).
			DateTime nowUtc = DateTime.UtcNow;
			DateTime windowEnd = nowUtc.AddDays(7);

			// Fuel listed before Not ALTN — same ordering used in the NOTAM Report tab's
			// Section() sub-headers, so the two surfaces read consistently.
			List<string> impactOrder = new List<string> { "A", "N", "C", "F", "D" };
			StringBuilder body = new StringBuilder();

			body.Append("<p class=\"introLine\">Below is a summary of operational impacts on the network over the next 7 days, cross-referenced against the flight schedule.</p>");

			if (flights.Count == 0)
				body.Append(
					"<div class=\"warnBanner\">" +
					"<span class=\"warnIcon\">&#9888;</span>" +
					"Flight Schedule is empty — no flights to cross-reference. Load the <b>Flight Schedule</b> tab first, then come back here." +
					"</div>");

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE Status='K' ORDER BY location", conn).ExecuteReader();

			int ordLocation  = reader.GetOrdinal("location");
			int ordKey       = reader.GetOrdinal("key");
			int ordAll       = reader.GetOrdinal("all");
			int ordStartdate = reader.GetOrdinal("startdate");
			int ordEnddate   = reader.GetOrdinal("enddate");
			int ordImpact    = reader.GetOrdinal("Impact");
			int ordRemark    = reader.GetOrdinal("Remark");

			// Grouped by impact code first, so each section's HTML can be assembled once
			// all matching NOTAMs for that code are known (needed for the count badge).
			Dictionary<string, List<string>> cardsByImpact = new Dictionary<string, List<string>>();
			foreach (string code in impactOrder) cardsByImpact[code] = new List<string>();

			while (reader.Read())
			{
				string impact = reader.IsDBNull(ordImpact) ? "" : reader.GetString(ordImpact);
				if (!cardsByImpact.ContainsKey(impact)) continue;   // not one of A/N/C/D/F (blank, or legacy)

				string location = reader.IsDBNull(ordLocation) ? "" : reader.GetString(ordLocation);
				string key      = reader.IsDBNull(ordKey) ? "" : reader.GetString(ordKey);
				string all      = reader.IsDBNull(ordAll) ? "" : reader.GetString(ordAll);
				string startRaw = reader.IsDBNull(ordStartdate) ? "" : reader.GetString(ordStartdate);
				string endRaw   = reader.IsDBNull(ordEnddate) ? "" : reader.GetString(ordEnddate);
				string remark   = reader.IsDBNull(ordRemark) ? "" : reader.GetString(ordRemark);
				if (location == "") continue;

				DateTime notamStart, notamEnd;
				if (!TryParseNotamDate(startRaw, out notamStart) || !TryParseNotamDate(endRaw, out notamEnd)) continue;

				List<string> matches = new List<string>();
				if (impact == "D")
				{
					// Not ALTN shows every Kept, D-classified NOTAM active at some point in
					// the next 7 days — full stop, no reference to flights at all. Whether an
					// airport can be used as an alternate isn't a function of which flights
					// happen to touch it as origin/destination, unlike every other impact
					// code, so there's no flight-matching step (and no IATA lookup) here.
					if (notamEnd < nowUtc || notamStart > windowEnd) continue;
				}
				else
				{
					// FlightSchedule.Origin/Dest are IATA codes (webservice's iataID, and the
					// CSV's DEP/ARR columns) while the NOTAM's location is ICAO — comparing
					// them directly never matched. Convert the NOTAM's station to IATA here.
					string iata = GetIATA(location);
					if (iata == "") continue;   // station not in Stations_ICAO_IATA — no IATA to match flights against

					foreach (FsFlight f in flights)
					{
						if (f.Origin == iata && f.HasStd && f.Std <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, f.Std))
							matches.Add(f.Callsign + " " + f.Origin + "-" + f.Dest + " — origin — STD " + FormatUtc(f.Std) + "Z");
						if (f.Dest == iata && f.HasSta && f.Sta <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, f.Sta))
							matches.Add(f.Callsign + " " + f.Origin + "-" + f.Dest + " — destination — STA " + FormatUtc(f.Sta) + "Z");
					}
					if (matches.Count == 0) continue;   // no conflict — nothing to show for this NOTAM
				}

				cardsByImpact[impact].Add(BuildConflictCardHtml(location, key, all, remark, matches, rasterDiagrams, cidImages, inlineImages));
			}
			conn.Close();

			foreach (string code in impactOrder)
			{
				Color c = ImpactColor(code);
				string hex = ColorTranslator.ToHtml(c);
				string label = ImpactLabel(code);
				List<string> cards = cardsByImpact[code];

				body.Append("<div class=\"sectionHeader\" style=\"border-left-color:" + hex + ";background:" + hex + "22" +
					(cards.Count == 0 ? ";opacity:.6" : "") + "\">" +
					"<span class=\"dot\" style=\"background:" + hex + "\"></span>" +
					"<span class=\"sectionTitle\">" + label + "</span>" +
					"<span class=\"count\" style=\"background:" + hex + "\">" + cards.Count + " conflit" + (cards.Count == 1 ? "" : "s") + "</span>" +
					"</div>");
				foreach (string card in cards) body.Append(card);
			}

			string html =
				"<html xmlns:v=\"urn:schemas-microsoft-com:vml\"><head><style>" +
				"v\\:*{behavior:url(#default#VML)}" +
				"body{margin:0;padding:16px;font-family:'Segoe UI',Arial,sans-serif;background:#fff;color:#222}" +
				".sectionHeader{position:relative;display:block;padding:10px 14px;border-left:4px solid;border-radius:6px;margin:0 0 8px 0}" +
				".dot{display:inline-block;width:10px;height:10px;border-radius:5px;margin-right:8px}" +
				".sectionTitle{font-size:15px;font-weight:bold;color:#222}" +
				".introLine{font-size:13.5px;color:#455a64;margin:0 0 16px 0}" +
				// Absolute rather than float:right — in the IE7-mode WebBrowser, a floated
				// span after inline content doesn't get cleared by the section header (whose
				// height then collapses around it), so the badge visually escapes its own
				// header and overlaps the airport card below instead of sitting flush right.
				".count{position:absolute;top:10px;right:14px;color:#fff;font-size:12px;padding:2px 10px;border-radius:10px}" +
				".card{border:1px solid #cfd8dc;border-radius:8px;overflow:hidden;margin:0 0 18px 0}" +
				".ahead{background:#263238;padding:14px 18px;position:relative}" +
				".icao{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".sub{font-size:13px;color:#78909c;margin-top:2px}" +
				".apname{font-size:13px;color:#90a4ae;margin-top:1px}" +
				".blk{font-size:12px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:6px 12px;margin-right:8px;vertical-align:top}" +
				".rwytable{margin-top:8px}" +
				".rwyline{white-space:nowrap;line-height:1.7}" +
				".diagram{position:absolute;top:10px;right:60px}" +
				".body{padding:12px 18px}" +
				".flightChip{display:inline-block;background:#fbe9e7;color:#4e342e;font-size:12px;padding:5px 10px;border-radius:6px;margin:0 8px 8px 0}" +
				".remark{font-size:12px;color:#455a64;margin:0 0 8px 0}" +
				".notamkey{font-size:12px;color:#607d8b;font-weight:bold;margin:0 0 4px 0}" +
				".notamtext{background:#f5f5f5;border-radius:6px;padding:10px 12px;font-family:'Courier New',monospace;font-size:12.5px;white-space:pre-wrap;line-height:1.6}" +
				".warnBanner{background:#fff3e0;color:#7a4a00;border:1px solid #ffcc80;border-radius:6px;padding:10px 14px;margin:0 0 16px 0;font-size:13px}" +
				".warnIcon{margin-right:8px}" +
				"</style></head><body>" + body + "</body></html>";

			return html;
		}

		private bool Overlaps(DateTime notamStart, DateTime notamEnd, DateTime flightTime)
		{
			DateTime winStart = flightTime.AddHours(-_conflictWindowHours);
			DateTime winEnd   = flightTime.AddHours(_conflictWindowHours);
			return notamStart <= winEnd && notamEnd >= winStart;
		}

		private static readonly string[] _dayTokens = { "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN" };
		private static readonly DayOfWeek[] _dayOfWeekByToken =
			{ DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday };

		private static int DayTokenIndex(string token)
		{
			for (int i = 0; i < _dayTokens.Length; i++) if (_dayTokens[i] == token.ToUpper()) return i;
			return -1;
		}

		private struct ScheduleEntry
		{
			public bool Daily;
			public HashSet<DayOfWeek> Days;
			public int StartMin, EndMin;   // minutes since local midnight, from HHMM
		}

		private static readonly Regex _dayItemRe = new Regex(
			@"\b(MON|TUE|WED|THU|FRI|SAT|SUN)\b(?:\s*-\s*\b(MON|TUE|WED|THU|FRI|SAT|SUN)\b)?\s+(\d{4})\s*-\s*(\d{4})",
			RegexOptions.IgnoreCase);
		private static readonly Regex _dailyRe = new Regex(@"\bDAILY\b\s+(\d{4})\s*-\s*(\d{4})", RegexOptions.IgnoreCase);

		// Extracts recurring D)-item closure schedules ("DAILY 2230-0330", "MON 0500-2359,
		// TUE 0000-0100", "MON-FRI 0600-1800", ...) out of the NOTAM's raw free text, and
		// expands them into concrete UTC occurrence intervals across [notamStart, notamEnd].
		// Returns null when nothing recognizable is found — the caller then falls back to
		// treating the whole validity window as active (today's behavior), so an exotic or
		// unparseable schedule never silently hides a real conflict.
		private static List<Tuple<DateTime, DateTime>> ParseNotamActiveWindows(string text, DateTime notamStart, DateTime notamEnd)
		{
			text = text ?? "";
			List<ScheduleEntry> entries = new List<ScheduleEntry>();

			MatchCollection dayMatches = _dayItemRe.Matches(text);
			if (dayMatches.Count > 0)
			{
				foreach (Match m in dayMatches)
				{
					int fromIdx = DayTokenIndex(m.Groups[1].Value);
					int toIdx = m.Groups[2].Success ? DayTokenIndex(m.Groups[2].Value) : fromIdx;
					if (fromIdx < 0 || toIdx < 0 || toIdx < fromIdx) continue;   // wraparound ranges (e.g. FRI-MON) unsupported — skip this entry
					int startMin, endMin;
					if (!TryHhmm(m.Groups[3].Value, out startMin) || !TryHhmm(m.Groups[4].Value, out endMin)) continue;

					HashSet<DayOfWeek> days = new HashSet<DayOfWeek>();
					for (int i = fromIdx; i <= toIdx; i++) days.Add(_dayOfWeekByToken[i]);
					entries.Add(new ScheduleEntry { Daily = false, Days = days, StartMin = startMin, EndMin = endMin });
				}
			}
			else
			{
				Match d = _dailyRe.Match(text);
				if (d.Success)
				{
					int startMin, endMin;
					if (TryHhmm(d.Groups[1].Value, out startMin) && TryHhmm(d.Groups[2].Value, out endMin))
						entries.Add(new ScheduleEntry { Daily = true, Days = null, StartMin = startMin, EndMin = endMin });
				}
			}

			if (entries.Count == 0) return null;

			List<Tuple<DateTime, DateTime>> occurrences = new List<Tuple<DateTime, DateTime>>();
			for (DateTime d = notamStart.Date; d <= notamEnd.Date; d = d.AddDays(1))
			{
				foreach (ScheduleEntry entry in entries)
				{
					if (!entry.Daily && !entry.Days.Contains(d.DayOfWeek)) continue;

					DateTime occStart = d.AddMinutes(entry.StartMin);
					DateTime occEnd = entry.EndMin > entry.StartMin ? d.AddMinutes(entry.EndMin) : d.AddDays(1).AddMinutes(entry.EndMin);

					DateTime clippedStart = occStart < notamStart ? notamStart : occStart;
					DateTime clippedEnd = occEnd > notamEnd ? notamEnd : occEnd;
					if (clippedStart < clippedEnd) occurrences.Add(new Tuple<DateTime, DateTime>(clippedStart, clippedEnd));
				}
			}
			return occurrences;
		}

		private static bool TryHhmm(string hhmm, out int minutesSinceMidnight)
		{
			minutesSinceMidnight = 0;
			int hour, minute;
			if (hhmm.Length != 4 || !int.TryParse(hhmm.Substring(0, 2), out hour) || !int.TryParse(hhmm.Substring(2, 2), out minute)) return false;
			if (hour > 24 || minute > 59) return false;
			minutesSinceMidnight = hour * 60 + minute;
			return true;
		}

		// Flight-time-vs-NOTAM match that honors a recurring D)-item schedule when one can be
		// recognized in the NOTAM text, instead of always treating the whole validity window
		// as active — fixes false conflicts like a Wednesday flight against a NOTAM only
		// closing "MON 0500-2359, TUE 0000-0100".
		private bool OverlapsSchedule(DateTime notamStart, DateTime notamEnd, string notamText, DateTime flightTime)
		{
			List<Tuple<DateTime, DateTime>> windows = ParseNotamActiveWindows(notamText, notamStart, notamEnd);
			if (windows == null) return Overlaps(notamStart, notamEnd, flightTime);
			foreach (Tuple<DateTime, DateTime> w in windows) if (Overlaps(w.Item1, w.Item2, flightTime)) return true;
			return false;
		}

		private static string FormatUtc(DateTime dt) { return dt.ToString("dd/MM HH:mm"); }

		// NOTAM start/enddate are actually stored ISO-style with a "T" separator (built in
		// GetXML as "yyyy-MM-ddTHH:mm:ss.000Z") — ParseExact against a space-separated
		// format silently failed for every row, so no NOTAM ever reached the overlap test
		// regardless of window size. Extracted positionally instead (same style as
		// dateTransformation(), which only reads the date portion and never cared whether
		// position 10 was 'T' or a space) so this is resilient to either separator.
		private static bool TryParseNotamDate(string raw, out DateTime result)
		{
			result = DateTime.MinValue;
			raw = (raw ?? "").Trim();
			if (raw.Length < 16) return false;
			try
			{
				int year   = int.Parse(raw.Substring(0, 4));
				int month  = int.Parse(raw.Substring(5, 2));
				int day    = int.Parse(raw.Substring(8, 2));
				int hour   = int.Parse(raw.Substring(11, 2));
				int minute = int.Parse(raw.Substring(14, 2));
				result = new DateTime(year, month, day, hour, minute, 0);
				return true;
			}
			catch { return false; }
		}

		private List<FsFlight> LoadFsFlights()
		{
			List<FsFlight> list = new List<FsFlight>();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT Callsign, Reg, Origin, Dest, STD, STA FROM FlightSchedule", conn).ExecuteReader();
			while (reader.Read())
			{
				FsFlight f = new FsFlight();
				f.Callsign = reader.IsDBNull(0) ? "" : reader.GetString(0);
				f.Reg      = reader.IsDBNull(1) ? "" : reader.GetString(1);
				f.Origin   = reader.IsDBNull(2) ? "" : reader.GetString(2);
				f.Dest     = reader.IsDBNull(3) ? "" : reader.GetString(3);
				string stdRaw = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string staRaw = reader.IsDBNull(5) ? "" : reader.GetString(5);
				f.HasStd = DateTime.TryParse(stdRaw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out f.Std);
				f.HasSta = DateTime.TryParse(staRaw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out f.Sta);
				list.Add(f);
			}
			conn.Close();
			return list;
		}

		// Airport card (ICAO/IATA/name/RWY table/diagram) + matching-flights chips + full
		// NOTAM text, as one <div class="card">. A leaner, self-contained twin of
		// BuildAirportCardHtml (which returns a whole standalone document sized for the
		// Filter tab's WebBrowser) since this fragment gets concatenated into one big
		// Conflict report page instead of living in its own WebBrowser control.
		private string BuildConflictCardHtml(string AP, string key, string notamText, string remark, List<string> matches,
			bool rasterDiagram, bool cidImages, Dictionary<string, string> inlineImages)
		{
			string flightChips = "";
			foreach (string m in matches) flightChips += "<span class=\"flightChip\">" + m + "</span>";

			string remarkLine = remark != "" ? "<div class=\"remark\">&#9654; " + remark.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string keyLine = key != "" ? "<div class=\"notamkey\">" + key.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";

			return
				"<div class=\"card\">" +
				BuildAirportHeaderHtml(AP, rasterDiagram, cidImages, inlineImages) +
				"<div class=\"body\">" +
				flightChips +
				remarkLine +
				keyLine +
				"<div class=\"notamtext\">" + notamText.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" +
				"</div>" +
				"</div>";
		}

		// The airport header block (ICAO/IATA/name/RWY table/diagram, i.e. the ".ahead" div)
		// shared between the Conflict card above and the per-airport detail sections the NOTAM
		// Report tab links a NOTAM key into (MainForm.Reports.cs) — pulled out so both stay
		// backed by the same RWY/geo-loading logic instead of two copies drifting apart.
		//
		// rasterDiagram/cidImages/inlineImages: see BuildConflictReportHtml's header comment.
		// inlineImages may be null when cidImages is false (data-URI/VML paths never touch it).
		public string BuildAirportHeaderHtml(string AP, bool rasterDiagram, bool cidImages, Dictionary<string, string> inlineImages)
		{
			string iata = GetIATA(AP);
			string name = GetAirportName(AP);

			string RWYs = "";
			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			while (OCCreader.Read()) if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			string[] rwyLines = RWYs.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
			List<string> rwyClean = new List<string>();
			foreach (string rl in rwyLines) if (rl.Trim() != "") rwyClean.Add(rl.Trim());

			List<RwyGeo> geo = LoadRwyGeo(AP);
			string rwySvg = rasterDiagram
				? BuildRwyDiagramImageTag(AP, geo, rwyClean, cidImages, inlineImages)
				: (HasGeo(geo) ? BuildRwySvgGeo(geo) : BuildRwySvg(rwyClean));

			string iataLine = (iata != "" && iata != AP) ? "<div class=\"sub\">IATA: " + iata + "</div>" : "";
			string nameLine = name != "" ? "<div class=\"apname\">" + name.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string leftCol = "", rightCol = "";
			for (int i = 0; i < rwyClean.Count; i++)
			{
				string cell = "<div class=\"rwyline\">" + rwyClean[i].Replace("&", "&amp;").Replace("<", "&lt;") + "</div>";
				if (i % 2 == 0) leftCol += cell; else rightCol += cell;
			}

			return
				"<div class=\"ahead\">" +
				"<div class=\"diagram\">" + rwySvg + "</div>" +
				"<div class=\"icao\">" + AP + "</div>" +
				iataLine + nameLine +
				"<table class=\"rwytable\" cellspacing=\"0\" cellpadding=\"0\"><tr>" +
				"<td class=\"blk\">" + leftCol + "</td><td class=\"blk\">" + rightCol + "</td>" +
				"</tr></table>" +
				"</div>";
		}

		// Renders the runway diagram as an <img> instead of VML (BuildRwyImage/BuildRwyImageGeo,
		// MainForm.NotamFilter.cs). cidImages=false embeds a data-URI (fine for wkhtmltopdf);
		// cidImages=true writes the PNG to a temp file once per airport and returns a
		// <img src="cid:..."> reference — the form Outlook's Word engine actually renders in an
		// HTML email body — recording (cid -> temp path) in inlineImages so MainForm.Email.cs
		// can attach it after the body is built.
		private static string BuildRwyDiagramImageTag(string AP, List<RwyGeo> geo, List<string> rwyClean, bool cidImages, Dictionary<string, string> inlineImages)
		{
			byte[] png = HasGeo(geo) ? BuildRwyImageGeo(geo) : BuildRwyImage(rwyClean);
			if (png == null || png.Length == 0) return "";

			if (!cidImages)
				return "<img width=\"130\" height=\"110\" src=\"data:image/png;base64," + Convert.ToBase64String(png) + "\">";

			string cid = "rwy_" + AP;
			if (!inlineImages.ContainsKey(cid))
			{
				string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), cid + ".png");
				System.IO.File.WriteAllBytes(tempPath, png);
				inlineImages[cid] = tempPath;
			}
			return "<img width=\"130\" height=\"110\" src=\"cid:" + cid + "\">";
		}
	}
}
