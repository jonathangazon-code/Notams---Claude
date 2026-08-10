using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace ICAO_CSV
{
	// Lets the Conflict tab's flight-match checkboxes (native <input type="checkbox">,
	// live WebBrowser only — never used in the email/PDF render) call back into C# to
	// dismiss a false-positive match or remove a manually-forced one, same ObjectForScripting
	// idiom as AviobookScriptBridge (MainForm.AipSup.cs) / ReportScriptBridge (MainForm.Reports.cs).
	[ComVisible(true)]
	public class ConflictScriptBridge
	{
		private MainForm _form;
		public ConflictScriptBridge(MainForm form) { _form = form; }
		public void DismissMatch(string notamKey, int fltlegId) { _form.DismissConflictMatch(notamKey, fltlegId); }
		public void RemoveManualConflict(int fltlegId, string notamKey) { _form.RemoveManualConflict(fltlegId, notamKey); }
	}

	public partial class MainForm
	{
		// One flight-schedule row, trimmed to what the conflict scan needs.
		private struct FsFlight
		{
			public int FltlegID;
			public string Callsign, Reg, Origin, Dest;
			public DateTime Std, Sta;
			public bool HasStd, HasSta;
			public string Alt1, Alt2;
			public int FlightTimeMin, Alt1TimeMin, Alt2TimeMin;
		}

		// One flight match rendered as a chip on a NOTAM card — Text is the display line,
		// FltlegId identifies which flight it is (for the dismiss/remove checkbox), and
		// Manual distinguishes a dispatcher-forced conflict (ManualConflicts table, removed
		// outright on uncheck) from an automatically-detected one (dismissed into
		// ConflictDismissals on uncheck, which only suppresses it — the underlying match
		// keeps being detected but is filtered back out on every future render).
		private struct ConflictMatch
		{
			public string Text;
			public int FltlegId;
			public bool Manual;
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
			Web_Conflict.ObjectForScripting = new ConflictScriptBridge(this);
			Web_Conflict.DocumentText = BuildConflictReportHtml(false, false, out unused);
		}

		// Called from the Conflict tab's flight-match checkbox (ConflictScriptBridge) when the
		// dispatcher unticks an automatically-detected match they consider a false positive.
		// Gated by EnsureWriterOrWarn like every other write path; either way the tab is
		// re-rendered afterward, which is also what makes a blocked Reader's checkbox visually
		// snap back to ticked (the dismissal never persisted, so the next render re-detects it).
		public void DismissConflictMatch(string notamKey, int fltlegId)
		{
			if (EnsureWriterOrWarn())
			{
				try { InsertConflictDismissal(notamKey, fltlegId); } catch { }
			}
			Build_Conflict_Report();
		}

		// Called from a manually-forced match's checkbox — unlike a dismissal, this deletes the
		// ManualConflicts row outright (there's nothing to "suppress", it's the dispatcher's own
		// record), so the flight's Flight Schedule "Force Conflict" button also reverts to unassigned.
		public void RemoveManualConflict(int fltlegId, string notamKey)
		{
			if (EnsureWriterOrWarn())
			{
				try { DeleteManualConflict(fltlegId); } catch { }
			}
			Build_Conflict_Report();
		}

		private void EnsureConflictDismissalsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE ConflictDismissals ([ID] AUTOINCREMENT PRIMARY KEY, [NotamKey] TEXT(50), [FltlegID] LONG, [DismissedAt] TEXT(25))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { }
		}

		private void InsertConflictDismissal(string notamKey, int fltlegId)
		{
			EnsureConflictDismissalsTable();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand ins = new OleDbCommand("INSERT INTO ConflictDismissals ([NotamKey],[FltlegID],[DismissedAt]) VALUES (?,?,?)", conn);
			ins.Parameters.AddWithValue("?", notamKey);
			ins.Parameters.AddWithValue("?", fltlegId);
			ins.Parameters.AddWithValue("?", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
			ins.ExecuteNonQuery();
			conn.Close();
		}

		// Keyed by "NotamKey|FltlegID" — checked before adding any automatically-detected
		// match (Origin/Dest/Alt1/Alt2) so a dismissed one simply never reappears.
		private HashSet<string> LoadDismissedConflicts()
		{
			HashSet<string> result = new HashSet<string>();
			EnsureConflictDismissalsTable();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader reader = new OleDbCommand("SELECT NotamKey, FltlegID FROM ConflictDismissals", conn).ExecuteReader();
				while (reader.Read())
				{
					string notamKey = reader.IsDBNull(0) ? "" : reader.GetString(0);
					int fltlegId = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetValue(1));
					result.Add(notamKey + "|" + fltlegId);
				}
				conn.Close();
			}
			catch { }
			return result;
		}

		private void EnsureManualConflictsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE ManualConflicts ([ID] AUTOINCREMENT PRIMARY KEY, [FltlegID] LONG, [NotamKey] TEXT(50), [CreatedAt] TEXT(25))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { }
		}

		// One active manual assignment per flight — replaces any prior row for the same FltlegID.
		public void AssignManualConflict(int fltlegId, string notamKey)
		{
			EnsureManualConflictsTable();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM ManualConflicts WHERE FltlegID=?", conn);
			del.Parameters.AddWithValue("?", fltlegId);
			del.ExecuteNonQuery();
			OleDbCommand ins = new OleDbCommand("INSERT INTO ManualConflicts ([FltlegID],[NotamKey],[CreatedAt]) VALUES (?,?,?)", conn);
			ins.Parameters.AddWithValue("?", fltlegId);
			ins.Parameters.AddWithValue("?", notamKey);
			ins.Parameters.AddWithValue("?", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
			ins.ExecuteNonQuery();
			conn.Close();
		}

		public void DeleteManualConflict(int fltlegId)
		{
			EnsureManualConflictsTable();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM ManualConflicts WHERE FltlegID=?", conn);
			del.Parameters.AddWithValue("?", fltlegId);
			del.ExecuteNonQuery();
			conn.Close();
		}

		// FltlegID -> assigned NotamKey ("" if none), for the Flight Schedule tab's "Force
		// Conflict" button text (MainForm.FlightSchedule.cs).
		public Dictionary<int, string> LoadManualConflictsByFltlegId()
		{
			Dictionary<int, string> result = new Dictionary<int, string>();
			EnsureManualConflictsTable();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader reader = new OleDbCommand("SELECT FltlegID, NotamKey FROM ManualConflicts", conn).ExecuteReader();
				while (reader.Read())
				{
					int fltlegId = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
					result[fltlegId] = reader.IsDBNull(1) ? "" : reader.GetString(1);
				}
				conn.Close();
			}
			catch { }
			return result;
		}

		// NotamKey -> list of manually-forced FltlegIDs, for BuildConflictReportHtml.
		private Dictionary<string, List<int>> LoadManualConflictsByNotamKey()
		{
			Dictionary<string, List<int>> result = new Dictionary<string, List<int>>();
			EnsureManualConflictsTable();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader reader = new OleDbCommand("SELECT FltlegID, NotamKey FROM ManualConflicts", conn).ExecuteReader();
				while (reader.Read())
				{
					int fltlegId = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
					string notamKey = reader.IsDBNull(1) ? "" : reader.GetString(1);
					if (notamKey == "") continue;
					List<int> list;
					if (!result.TryGetValue(notamKey, out list)) result[notamKey] = list = new List<int>();
					list.Add(fltlegId);
				}
				conn.Close();
			}
			catch { }
			return result;
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
			Dictionary<int, FsFlight> flightsById = new Dictionary<int, FsFlight>();
			foreach (FsFlight f in flights) flightsById[f.FltlegID] = f;
			HashSet<string> dismissed = LoadDismissedConflicts();
			Dictionary<string, List<int>> manualByNotamKey = LoadManualConflictsByNotamKey();
			// Live tab only (rasterDiagrams=false) gets interactive dismiss/remove checkboxes on
			// each flight chip — the email/PDF render (rasterDiagrams=true) has no JS bridge and
			// stays plain text, same convention rasterDiagrams already uses everywhere else.
			bool interactive = !rasterDiagrams;

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

				List<ConflictMatch> matches = new List<ConflictMatch>();
				bool altnConflict = false;
				if (impact == "D")
				{
					// Not ALTN shows every Kept, D-classified NOTAM active at some point in
					// the next 7 days regardless of flights — whether an airport can be used as
					// an alternate isn't a function of origin/destination traffic. But if this
					// station is actually filed as a flight's diversion alternate (Alt1/Alt2,
					// from the briefing's AlternateFuel data), that's a real, time-specific
					// conflict worth calling out distinctly: estimated arrival at the alternate
					// is STD + main-leg flight time + alternate flight time, checked against the
					// NOTAM's active window with its own ± window (_altnConflictWindowHours,
					// independent of the origin/destination one).
					if (notamEnd < nowUtc || notamStart > windowEnd) continue;

					foreach (FsFlight f in flights)
					{
						if (!f.HasStd) continue;
						if (f.Alt1 == location && f.FlightTimeMin > 0 && f.Alt1TimeMin > 0 && !dismissed.Contains(key + "|" + f.FltlegID))
						{
							DateTime altArrival = f.Std.AddMinutes(f.FlightTimeMin + f.Alt1TimeMin);
							if (altArrival <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, altArrival, _altnConflictWindowHours))
							{
								matches.Add(new ConflictMatch { FltlegId = f.FltlegID, Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — alternate 1 — est. diversion arrival " + FormatUtc(altArrival) + "Z" });
								altnConflict = true;
							}
						}
						if (f.Alt2 == location && f.FlightTimeMin > 0 && f.Alt2TimeMin > 0 && !dismissed.Contains(key + "|" + f.FltlegID))
						{
							DateTime altArrival = f.Std.AddMinutes(f.FlightTimeMin + f.Alt2TimeMin);
							if (altArrival <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, altArrival, _altnConflictWindowHours))
							{
								matches.Add(new ConflictMatch { FltlegId = f.FltlegID, Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — alternate 2 — est. diversion arrival " + FormatUtc(altArrival) + "Z" });
								altnConflict = true;
							}
						}
					}
				}
				else
				{
					// FlightSchedule.Origin/Dest are IATA codes (webservice's iataID, and the
					// CSV's DEP/ARR columns) while the NOTAM's location is ICAO — comparing
					// them directly never matched. Convert the NOTAM's station to IATA here.
					string iata = GetIATA(location);
					if (iata != "")
					{
						foreach (FsFlight f in flights)
						{
							if (dismissed.Contains(key + "|" + f.FltlegID)) continue;
							if (f.Origin == iata && f.HasStd && f.Std <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, f.Std))
								matches.Add(new ConflictMatch { FltlegId = f.FltlegID, Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — origin — STD " + FormatUtc(f.Std) + "Z" });
							if (f.Dest == iata && f.HasSta && f.Sta <= windowEnd && OverlapsSchedule(notamStart, notamEnd, all, f.Sta))
								matches.Add(new ConflictMatch { FltlegId = f.FltlegID, Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — destination — STA " + FormatUtc(f.Sta) + "Z" });
						}
					}
					// station not in Stations_ICAO_IATA (iata=="") just means no automatic
					// match is possible — a manually-forced one can still apply below.
				}

				// Manually-forced flights (Flight Schedule tab's "Force Conflict") always apply,
				// regardless of impact code or whether the automatic matching found anything —
				// this is how a dispatcher surfaces a false negative the algorithm can't express.
				List<int> manualIds;
				bool hasManualMatch = false;
				if (manualByNotamKey.TryGetValue(key, out manualIds))
				{
					foreach (int fltlegId in manualIds)
					{
						FsFlight f;
						if (!flightsById.TryGetValue(fltlegId, out f)) continue;   // flight no longer in FlightSchedule (past/dropped)
						matches.Add(new ConflictMatch { FltlegId = fltlegId, Manual = true,
							Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — forced by dispatcher" });
						hasManualMatch = true;
					}
				}

				bool highlight = altnConflict || hasManualMatch;
				if (impact != "D" && matches.Count == 0) continue;   // no automatic or manual conflict — nothing to show for this NOTAM

				string cardHtml = BuildConflictCardHtml(location, key, all, remark, matches, rasterDiagrams, cidImages, inlineImages, highlight, interactive);
				if (highlight) cardsByImpact[impact].Insert(0, cardHtml);
				else cardsByImpact[impact].Add(cardHtml);
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
				".cardAlert{border:2px solid #c62828;box-shadow:0 0 0 1px #c62828}" +
				".flightChipAlert{background:#c62828;color:#fff;font-weight:bold}" +
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
				".flightChipManual{background:#ff8f00;color:#fff;font-weight:bold}" +
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
			return Overlaps(notamStart, notamEnd, flightTime, _conflictWindowHours);
		}

		private bool Overlaps(DateTime notamStart, DateTime notamEnd, DateTime flightTime, int windowHours)
		{
			DateTime winStart = flightTime.AddHours(-windowHours);
			DateTime winEnd   = flightTime.AddHours(windowHours);
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

		// Kind distinguishes the three recurring-schedule shapes a D)-item can use — a plain
		// day-of-week list (Weekly), a day-of-month/month+day(-range) list (MonthDay, e.g.
		// "11 2330-0245, 13 0001-0245" or "AUG 11-13 0530-1600"), or DAILY.
		private const int ScheduleKindWeekly = 0;
		private const int ScheduleKindMonthDay = 1;
		private const int ScheduleKindDaily = 2;

		private struct ScheduleEntry
		{
			public int Kind;
			public HashSet<DayOfWeek> Days;   // Kind=Weekly
			public int Month;                 // Kind=MonthDay, 0 = unspecified/any month
			public int Day1, Day2;            // Kind=MonthDay, inclusive day-of-month range
			public int StartMin, EndMin;      // minutes since local midnight, from HHMM
		}

		private static readonly string[] _monthTokens = { "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };

		private static int MonthTokenIndex(string token)
		{
			for (int i = 0; i < _monthTokens.Length; i++) if (_monthTokens[i] == token.ToUpper()) return i;
			return -1;
		}

		private static readonly Regex _dayItemRe = new Regex(
			@"\b(MON|TUE|WED|THU|FRI|SAT|SUN)\b(?:\s*-\s*\b(MON|TUE|WED|THU|FRI|SAT|SUN)\b)?\s+(\d{4})\s*-\s*(\d{4})",
			RegexOptions.IgnoreCase);
		private static readonly Regex _dailyRe = new Regex(@"\bDAILY\b\s+(\d{4})\s*-\s*(\d{4})", RegexOptions.IgnoreCase);
		// Day-of-month, optionally prefixed by a month abbreviation, optionally a same-month
		// day range (e.g. "11 2330-0245", "AUG 10 1100-1600", "AUG 11-13 0530-1600"). Can't
		// collide with _dayItemRe: that one requires a MON/TUE/... token where this one
		// requires a digit at the same position.
		private static readonly Regex _monthDayItemRe = new Regex(
			@"\b(?:(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+)?(\d{1,2})(?:\s*-\s*(\d{1,2}))?\s+(\d{4})\s*-\s*(\d{4})\b",
			RegexOptions.IgnoreCase);

		// Extracts recurring D)-item closure schedules ("DAILY 2230-0330", "MON 0500-2359,
		// TUE 0000-0100", "MON-FRI 0600-1800", "11 2330-0245, 13 0001-0245", "AUG 11-13
		// 0530-1600", ...) out of the NOTAM's raw free text, and expands them into concrete
		// UTC occurrence intervals across [notamStart, notamEnd]. Returns null when nothing
		// recognizable is found — the caller then falls back to treating the whole validity
		// window as active (today's behavior), so an exotic or unparseable schedule never
		// silently hides a real conflict.
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
					entries.Add(new ScheduleEntry { Kind = ScheduleKindWeekly, Days = days, StartMin = startMin, EndMin = endMin });
				}
			}
			else
			{
				MatchCollection monthDayMatches = _monthDayItemRe.Matches(text);
				if (monthDayMatches.Count > 0)
				{
					foreach (Match m in monthDayMatches)
					{
						string monthTok = m.Groups[1].Value;
						int month = monthTok == "" ? 0 : MonthTokenIndex(monthTok) + 1;
						if (monthTok != "" && month <= 0) continue;   // shouldn't happen given the regex, but stay safe

						int day1, day2;
						if (!int.TryParse(m.Groups[2].Value, out day1)) continue;
						day2 = m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : day1;
						if (day1 < 1 || day1 > 31 || day2 < day1) continue;   // wraparound day ranges unsupported — skip this entry

						int startMin, endMin;
						if (!TryHhmm(m.Groups[4].Value, out startMin) || !TryHhmm(m.Groups[5].Value, out endMin)) continue;

						entries.Add(new ScheduleEntry { Kind = ScheduleKindMonthDay, Month = month, Day1 = day1, Day2 = day2, StartMin = startMin, EndMin = endMin });
					}
				}
				else
				{
					Match d = _dailyRe.Match(text);
					if (d.Success)
					{
						int startMin, endMin;
						if (TryHhmm(d.Groups[1].Value, out startMin) && TryHhmm(d.Groups[2].Value, out endMin))
							entries.Add(new ScheduleEntry { Kind = ScheduleKindDaily, StartMin = startMin, EndMin = endMin });
					}
				}
			}

			if (entries.Count == 0) return null;

			List<Tuple<DateTime, DateTime>> occurrences = new List<Tuple<DateTime, DateTime>>();
			for (DateTime d = notamStart.Date; d <= notamEnd.Date; d = d.AddDays(1))
			{
				foreach (ScheduleEntry entry in entries)
				{
					if (!EntryMatchesDate(entry, d)) continue;

					DateTime occStart = d.AddMinutes(entry.StartMin);
					DateTime occEnd = entry.EndMin > entry.StartMin ? d.AddMinutes(entry.EndMin) : d.AddDays(1).AddMinutes(entry.EndMin);

					DateTime clippedStart = occStart < notamStart ? notamStart : occStart;
					DateTime clippedEnd = occEnd > notamEnd ? notamEnd : occEnd;
					if (clippedStart < clippedEnd) occurrences.Add(new Tuple<DateTime, DateTime>(clippedStart, clippedEnd));
				}
			}
			return occurrences;
		}

		private static bool EntryMatchesDate(ScheduleEntry entry, DateTime d)
		{
			switch (entry.Kind)
			{
				case ScheduleKindWeekly: return entry.Days.Contains(d.DayOfWeek);
				case ScheduleKindMonthDay: return (entry.Month == 0 || d.Month == entry.Month) && d.Day >= entry.Day1 && d.Day <= entry.Day2;
				default: return true;   // ScheduleKindDaily
			}
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
			return OverlapsSchedule(notamStart, notamEnd, notamText, flightTime, _conflictWindowHours);
		}

		private bool OverlapsSchedule(DateTime notamStart, DateTime notamEnd, string notamText, DateTime flightTime, int windowHours)
		{
			List<Tuple<DateTime, DateTime>> windows = ParseNotamActiveWindows(notamText, notamStart, notamEnd);
			if (windows == null) return Overlaps(notamStart, notamEnd, flightTime, windowHours);
			foreach (Tuple<DateTime, DateTime> w in windows) if (Overlaps(w.Item1, w.Item2, flightTime, windowHours)) return true;
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
			OleDbDataReader reader = new OleDbCommand(
				"SELECT FltlegID, Callsign, Reg, Origin, Dest, STD, STA, Alt1, Alt2, FlightTimeMin, Alt1TimeMin, Alt2TimeMin FROM FlightSchedule", conn).ExecuteReader();
			while (reader.Read())
			{
				FsFlight f = new FsFlight();
				f.FltlegID = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
				f.Callsign = reader.IsDBNull(1) ? "" : reader.GetString(1);
				f.Reg      = reader.IsDBNull(2) ? "" : reader.GetString(2);
				f.Origin   = reader.IsDBNull(3) ? "" : reader.GetString(3);
				f.Dest     = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string stdRaw = reader.IsDBNull(5) ? "" : reader.GetString(5);
				string staRaw = reader.IsDBNull(6) ? "" : reader.GetString(6);
				f.HasStd = DateTime.TryParse(stdRaw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out f.Std);
				f.HasSta = DateTime.TryParse(staRaw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out f.Sta);
				f.Alt1 = reader.IsDBNull(7) ? "" : reader.GetString(7);
				f.Alt2 = reader.IsDBNull(8) ? "" : reader.GetString(8);
				f.FlightTimeMin = reader.IsDBNull(9)  ? 0 : Convert.ToInt32(reader.GetValue(9));
				f.Alt1TimeMin   = reader.IsDBNull(10) ? 0 : Convert.ToInt32(reader.GetValue(10));
				f.Alt2TimeMin   = reader.IsDBNull(11) ? 0 : Convert.ToInt32(reader.GetValue(11));
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
		private string BuildConflictCardHtml(string AP, string key, string notamText, string remark, List<ConflictMatch> matches,
			bool rasterDiagram, bool cidImages, Dictionary<string, string> inlineImages, bool highlight, bool interactive)
		{
			string flightChips = "";
			foreach (ConflictMatch m in matches)
			{
				string chipClass = "flightChip" + (m.Manual ? " flightChipManual" : (highlight ? " flightChipAlert" : ""));
				string checkbox = "";
				if (interactive)
				{
					string jsCall = m.Manual
						? "window.external.RemoveManualConflict(" + m.FltlegId + ",'" + key + "')"
						: "window.external.DismissMatch('" + key + "'," + m.FltlegId + ")";
					checkbox = "<input type=\"checkbox\" checked=\"checked\" onclick=\"" + jsCall + "\"> ";
				}
				flightChips += "<span class=\"" + chipClass + "\">" + checkbox + m.Text + "</span>";
			}

			string remarkLine = remark != "" ? "<div class=\"remark\">&#9654; " + remark.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string keyLine = key != "" ? "<div class=\"notamkey\">" + key.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";

			// A Not-ALTN NOTAM that's a real diversion-time conflict against a flight's filed
			// alternate is a higher-severity finding than the section's default "airport isn't
			// usable as an alternate at all, regardless of any flight" listing — highlighted red
			// and (by the caller, cardsByImpact.Insert(0, ...)) sorted to the top of its section.
			string cardClass = highlight ? "card cardAlert" : "card";

			return
				"<div class=\"" + cardClass + "\">" +
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
			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
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

			string iataLine = (iata != "" && iata != AP) ? "<div class=\"sub\">IATA: " + iata + "</div>" : "";
			string nameLine = name != "" ? "<div class=\"apname\">" + name.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string leftCol = "", rightCol = "";
			for (int i = 0; i < rwyClean.Count; i++)
			{
				string cell = "<div class=\"rwyline\">" + rwyClean[i].Replace("&", "&amp;").Replace("<", "&lt;") + "</div>";
				if (i % 2 == 0) leftCol += cell; else rightCol += cell;
			}
			string textBlock =
				"<div class=\"icao\">" + AP + "</div>" +
				iataLine + nameLine +
				"<table class=\"rwytable\" cellspacing=\"0\" cellpadding=\"0\"><tr>" +
				"<td class=\"blk\">" + leftCol + "</td><td class=\"blk\">" + rightCol + "</td>" +
				"</tr></table>";

			if (!rasterDiagram)
			{
				string rwySvg = HasGeo(geo) ? BuildRwySvgGeo(geo) : BuildRwySvg(rwyClean);
				return
					"<div class=\"ahead\">" +
					"<div class=\"diagram\">" + rwySvg + "</div>" +
					textBlock +
					"</div>";
			}

			// Outlook's Word rendering engine doesn't reliably honor "position:absolute;
			// right:..." (the VML/live-tab layout above) — the diagram ends up flush left
			// instead of on the right. A table-based two-column layout is the standard
			// Outlook-safe workaround, and renders identically in wkhtmltopdf too, so this path
			// is used for both the NOTAM Report PDF and the email.
			string img = BuildRwyDiagramImageTag(AP, geo, rwyClean, cidImages, inlineImages);
			return
				"<div class=\"ahead\">" +
				"<table class=\"aheadTable\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\"><tr>" +
				"<td valign=\"top\">" + textBlock + "</td>" +
				"<td valign=\"top\" align=\"right\" width=\"140\">" + img + "</td>" +
				"</tr></table>" +
				"</div>";
		}

		// Renders the runway diagram as an <img> instead of VML (BuildRwyImage/BuildRwyImageGeo,
		// MainForm.NotamFilter.cs). Always writes the PNG to a stable per-airport temp file
		// (deterministic name, overwritten each run — its content is a pure function of the
		// airport's RWY data, so reuse across runs is harmless) rather than embedding it as a
		// data-URI: wkhtmltopdf's (dated, Qt-WebKit-based) data-URI <img> support turned out to
		// be unreliable in practice (only the top few PNG rows would render), and a real
		// file:// reference is far more universally supported. cidImages=false (NOTAM Report
		// PDF export + its live WebBrowser preview) embeds a <img src="file:///..."> reference;
		// cidImages=true (email) embeds <img src="cid:..."> instead — the only form Outlook's
		// Word engine renders in an HTML body — recording (cid -> temp path) in inlineImages so
		// MainForm.Email.cs can attach it after the body is built.
		private static string BuildRwyDiagramImageTag(string AP, List<RwyGeo> geo, List<string> rwyClean, bool cidImages, Dictionary<string, string> inlineImages)
		{
			byte[] png = HasGeo(geo) ? BuildRwyImageGeo(geo) : BuildRwyImage(rwyClean);
			if (png == null || png.Length == 0) return "";

			string cid = "rwy_" + AP;
			string tempPath;
			if (inlineImages != null && inlineImages.ContainsKey(cid))
			{
				tempPath = inlineImages[cid];
			}
			else
			{
				tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), cid + ".png");
				System.IO.File.WriteAllBytes(tempPath, png);
				if (inlineImages != null) inlineImages[cid] = tempPath;
			}

			string src = cidImages ? "cid:" + cid : new Uri(tempPath).AbsoluteUri;
			return "<img width=\"130\" height=\"110\" src=\"" + src + "\">";
		}
	}
}
