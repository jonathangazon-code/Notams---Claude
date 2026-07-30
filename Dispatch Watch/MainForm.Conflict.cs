using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Text;

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
		// a conflict exists when the NOTAM's validity window overlaps the ±2h window
		// around a flight's STD (station used as Origin) or STA (station used as Dest).
		// One section per impact code (A/N/C/D/F, same severity order as
		// SuggestedSingleCode), each listing the airport card + matching flights + full
		// NOTAM text for every conflict found; SUP is out of scope for this tab.
		public void Build_Conflict_Report()
		{
			List<FsFlight> flights = LoadFsFlights();

			List<string> impactOrder = new List<string> { "A", "N", "C", "D", "F" };
			StringBuilder body = new StringBuilder();

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
				foreach (FsFlight f in flights)
				{
					if (f.Origin == location && f.HasStd && Overlaps(notamStart, notamEnd, f.Std))
						matches.Add(f.Callsign + " — origin — STD " + FormatUtc(f.Std) + "Z");
					if (f.Dest == location && f.HasSta && Overlaps(notamStart, notamEnd, f.Sta))
						matches.Add(f.Callsign + " — destination — STA " + FormatUtc(f.Sta) + "Z");
				}
				if (matches.Count == 0) continue;   // no conflict — nothing to show for this NOTAM

				cardsByImpact[impact].Add(BuildConflictCardHtml(location, key, all, remark, matches));
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
				".sectionHeader{display:block;padding:10px 14px;border-left:4px solid;border-radius:6px;margin:0 0 8px 0}" +
				".dot{display:inline-block;width:10px;height:10px;border-radius:5px;margin-right:8px}" +
				".sectionTitle{font-size:15px;font-weight:bold;color:#222}" +
				".count{float:right;color:#fff;font-size:12px;padding:2px 10px;border-radius:10px}" +
				".card{border:1px solid #cfd8dc;border-radius:8px;overflow:hidden;margin:0 0 18px 0}" +
				".ahead{background:#263238;padding:14px 18px;position:relative}" +
				".icao{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".sub{font-size:13px;color:#78909c;margin-top:2px}" +
				".apname{font-size:13px;color:#90a4ae;margin-top:1px}" +
				".blk{font-size:12px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:6px 12px;margin-right:8px;vertical-align:top}" +
				".rwytable{margin-top:8px}" +
				".rwyline{white-space:nowrap;line-height:1.7}" +
				".diagram{position:absolute;top:10px;right:16px}" +
				".body{padding:12px 18px}" +
				".flightChip{display:inline-block;background:#fbe9e7;color:#4e342e;font-size:12px;padding:5px 10px;border-radius:6px;margin:0 8px 8px 0}" +
				".remark{font-size:12px;color:#455a64;margin:0 0 8px 0}" +
				".notamkey{font-size:12px;color:#607d8b;font-weight:bold;margin:0 0 4px 0}" +
				".notamtext{background:#f5f5f5;border-radius:6px;padding:10px 12px;font-family:'Courier New',monospace;font-size:12.5px;white-space:pre-wrap;line-height:1.6}" +
				"</style></head><body>" + body + "</body></html>";

			Web_Conflict.DocumentText = html;
		}

		// Widened from the intended ±2h to ±12h for testing (no conflicts were showing up
		// with the tighter window) — revert to 2 once verified against real data.
		private const int ConflictWindowHours = 12;

		private bool Overlaps(DateTime notamStart, DateTime notamEnd, DateTime flightTime)
		{
			DateTime winStart = flightTime.AddHours(-ConflictWindowHours);
			DateTime winEnd   = flightTime.AddHours(ConflictWindowHours);
			return notamStart <= winEnd && notamEnd >= winStart;
		}

		private static string FormatUtc(DateTime dt) { return dt.ToString("dd/MM HH:mm"); }

		// NOTAM start/enddate are stored as "yyyy-MM-dd HH:mm..." (see dateTransformation) —
		// only the first 16 characters ("yyyy-MM-dd HH:mm") are the actual date/time.
		private static bool TryParseNotamDate(string raw, out DateTime result)
		{
			raw = (raw ?? "").Trim();
			if (raw.Length < 16) { result = DateTime.MinValue; return false; }
			return DateTime.TryParseExact(raw.Substring(0, 16), "yyyy-MM-dd HH:mm",
				CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
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
		private string BuildConflictCardHtml(string AP, string key, string notamText, string remark, List<string> matches)
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
			string rwySvg = HasGeo(geo) ? BuildRwySvgGeo(geo) : BuildRwySvg(rwyClean);

			string iataLine = (iata != "" && iata != AP) ? "<div class=\"sub\">IATA: " + iata + "</div>" : "";
			string nameLine = name != "" ? "<div class=\"apname\">" + name.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string leftCol = "", rightCol = "";
			for (int i = 0; i < rwyClean.Count; i++)
			{
				string cell = "<div class=\"rwyline\">" + rwyClean[i].Replace("&", "&amp;").Replace("<", "&lt;") + "</div>";
				if (i % 2 == 0) leftCol += cell; else rightCol += cell;
			}

			string flightChips = "";
			foreach (string m in matches) flightChips += "<span class=\"flightChip\">" + m + "</span>";

			string remarkLine = remark != "" ? "<div class=\"remark\">&#9654; " + remark.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";
			string keyLine = key != "" ? "<div class=\"notamkey\">" + key.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" : "";

			return
				"<div class=\"card\">" +
				"<div class=\"ahead\">" +
				"<div class=\"diagram\">" + rwySvg + "</div>" +
				"<div class=\"icao\">" + AP + "</div>" +
				iataLine + nameLine +
				"<table class=\"rwytable\" cellspacing=\"0\" cellpadding=\"0\"><tr>" +
				"<td class=\"blk\">" + leftCol + "</td><td class=\"blk\">" + rightCol + "</td>" +
				"</tr></table>" +
				"</div>" +
				"<div class=\"body\">" +
				flightChips +
				remarkLine +
				keyLine +
				"<div class=\"notamtext\">" + notamText.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>" +
				"</div>" +
				"</div>";
		}
	}
}
