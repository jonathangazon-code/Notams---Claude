using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	// TAF Analysis tab: downloads the FT (TAF) for every Airport List station from the fleet's
	// weather webservice, applies the same threshold/color-coding logic as the old standalone
	// "TAF analysis" app (TAF analysis/Dispatch/MainForm.cs, Btn_addTAFClick), and shows a
	// single shared, dispatcher-editable report — sent by email to its own recipient list,
	// independent of the Conflict tab's EmailRecipients.
	public partial class MainForm
	{
		// Threshold defaults match the legacy standalone app's hardcoded values. Persisted via
		// ArchiveConfig.xml (MainForm.Admin.cs's EnsureArchiveConfig/SaveArchiveConfig), but
		// edited from this tab's own top bar rather than the Admin tab, since these are the
		// dispatcher's day-to-day working values, not one-off endpoint config.
		private static int _tafVisCatIm              = 550;
		private static int _tafVisAdvisoryM          = 1000;
		private static int _tafCeilCatIHundredFt     = 2;
		private static int _tafCeilAdvisoryHundredFt = 4;
		private static int _tafWindCatIKt            = 44;
		private static int _tafWindAdvisoryKt        = 34;
		private static int _tafWindCatIMps           = 22;
		private static int _tafWindAdvisoryMps       = 17;

		private Panel _tafTopBar;
		private TextBox _tafVisCatIBox, _tafVisAdvBox, _tafCeilCatIBox, _tafCeilAdvBox,
			_tafWindKtCatIBox, _tafWindKtAdvBox, _tafWindMpsCatIBox, _tafWindMpsAdvBox;
		private bool _tafBodyLoaded;
		// Session-only (not persisted) — true once DownloadAndAnalyzeTafs() has actually
		// completed a run in this app session. TafReport.Html can be stale from a previous
		// session/dispatcher (or never generated on this station's DB copy at all), so
		// Btn_sendTafReportClick warns before sending if nobody has clicked "TAF Analysis"
		// since this instance started, even though a stored report might still exist.
		private bool _tafAnalysisRunThisSession;

		// ── TafRecipients: independent recipient list from the Conflict tab's EmailRecipients ──

		public void EnsureTafRecipientsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE TafRecipients ([Email] TEXT(255))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		List<string> LoadTafRecipients()
		{
			List<string> list = new List<string>();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Email FROM TafRecipients ORDER BY Email", conn).ExecuteReader();
				while (r.Read())
					if (!r.IsDBNull(0) && r.GetString(0).Trim() != "") list.Add(r.GetString(0).Trim());
				conn.Close();
			}
			catch { }
			return list;
		}

		// ── TafReport: single shared row holding the current (possibly dispatcher-edited) report ──

		public void EnsureTafReportTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE TafReport ([ID] LONG, [Html] MEMO, [TimeIssued] TEXT(50))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try
				{
					object count = new OleDbCommand("SELECT COUNT(*) FROM TafReport WHERE ID=1", conn).ExecuteScalar();
					if (Convert.ToInt32(count) == 0)
						new OleDbCommand("INSERT INTO TafReport ([ID],[Html],[TimeIssued]) VALUES (1,'','')", conn).ExecuteNonQuery();
				}
				catch { }
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		string LoadTafReportHtml(out string timeIssued)
		{
			timeIssued = "";
			string html = "";
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Html,TimeIssued FROM TafReport WHERE ID=1", conn).ExecuteReader();
				if (r.Read())
				{
					if (!r.IsDBNull(0)) html = r.GetString(0);
					if (!r.IsDBNull(1)) timeIssued = r.GetString(1);
				}
				conn.Close();
			}
			catch { }
			return html;
		}

		void SaveTafReportHtml(string html, string timeIssued)
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbCommand upd = new OleDbCommand("UPDATE TafReport SET Html=?, TimeIssued=? WHERE ID=1", conn);
				upd.Parameters.AddWithValue("?", html);
				upd.Parameters.AddWithValue("?", timeIssued);
				upd.ExecuteNonQuery();
				conn.Close();
			}
			catch { }
		}

		// ── TafStationFragments: snapshot of every red (CAT I)-flagged station from the last
		// analysis, keyed by ICAO — the input BuildTafCheckVsScheduleHtml needs to rebuild the
		// "Check vs schedule" section on demand (live tab render, Send) without re-downloading
		// or re-parsing the raw TAF text, which isn't stored anywhere. Only flagged stations are
		// kept (cleared and rewritten on every DownloadAndAnalyzeTafs() run) — a station with no
		// red fragment at all is simply absent from the table. ──

		public void EnsureTafStationFragmentsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try
				{
					new OleDbCommand(
						"CREATE TABLE TafStationFragments ([ICAO] TEXT(4), [VisCeil] MEMO, [Wind] MEMO, [TS] MEMO, [Snow] MEMO, " +
						"[VisCeilWindows] MEMO, [WindWindows] MEMO, [TSWindows] MEMO, [SnowWindows] MEMO)", conn)
						.ExecuteNonQuery();
				}
				catch
				{
					// Already exists — but may predate the *Windows columns (added after this
					// table's original CREATE TABLE shipped), in which case SaveTafStationFragments'
					// INSERT would fail on every single run (unknown column), get swallowed by its
					// own try/catch, and silently leave the table empty forever. Each ALTER is its
					// own try/catch since Jet/ACE has no "ADD COLUMN IF NOT EXISTS" — failure here
					// just means that particular column already exists.
					try { new OleDbCommand("ALTER TABLE TafStationFragments ADD COLUMN [VisCeilWindows] MEMO", conn).ExecuteNonQuery(); } catch { }
					try { new OleDbCommand("ALTER TABLE TafStationFragments ADD COLUMN [WindWindows] MEMO", conn).ExecuteNonQuery(); } catch { }
					try { new OleDbCommand("ALTER TABLE TafStationFragments ADD COLUMN [TSWindows] MEMO", conn).ExecuteNonQuery(); } catch { }
					try { new OleDbCommand("ALTER TABLE TafStationFragments ADD COLUMN [SnowWindows] MEMO", conn).ExecuteNonQuery(); } catch { }
				}
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		private void SaveTafStationFragments(Dictionary<string, TafFragments> fragByIcao)
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				new OleDbCommand("DELETE FROM TafStationFragments", conn).ExecuteNonQuery();
				foreach (KeyValuePair<string, TafFragments> kv in fragByIcao)
				{
					TafFragments f = kv.Value;
					// Only a CAT I (red)-flagged station is worth keeping — TS/Snow entries are
					// always red, Vis/Ceiling/Wind may be red, amber, or blue.
					bool anyRed = ContainsRed(f.VisCeil) || ContainsRed(f.Wind) || ContainsRed(f.TS) || ContainsRed(f.Snow);
					if (!anyRed) continue;

					// Per-row try/catch: one station's row failing to insert (e.g. a schema still
					// missing a column despite EnsureTafStationFragmentsTable's migration) must not
					// silently drop every other station too — the previous single try/catch around
					// the whole method did exactly that.
					try
					{
						OleDbCommand ins = new OleDbCommand(
							"INSERT INTO TafStationFragments ([ICAO],[VisCeil],[Wind],[TS],[Snow]," +
							"[VisCeilWindows],[WindWindows],[TSWindows],[SnowWindows]) VALUES (?,?,?,?,?,?,?,?,?)", conn);
						ins.Parameters.AddWithValue("?", kv.Key);
						ins.Parameters.AddWithValue("?", f.VisCeil ?? "");
						ins.Parameters.AddWithValue("?", f.Wind ?? "");
						ins.Parameters.AddWithValue("?", f.TS ?? "");
						ins.Parameters.AddWithValue("?", f.Snow ?? "");
						ins.Parameters.AddWithValue("?", SerializeWindows(f.VisCeilWindows));
						ins.Parameters.AddWithValue("?", SerializeWindows(f.WindWindows));
						ins.Parameters.AddWithValue("?", SerializeWindows(f.TSWindows));
						ins.Parameters.AddWithValue("?", SerializeWindows(f.SnowWindows));
						ins.ExecuteNonQuery();
					}
					catch { }
				}
				conn.Close();
			}
			catch { }
		}

		private Dictionary<string, TafFragments> LoadTafStationFragments()
		{
			Dictionary<string, TafFragments> result = new Dictionary<string, TafFragments>(StringComparer.OrdinalIgnoreCase);
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand(
					"SELECT ICAO,VisCeil,Wind,TS,Snow,VisCeilWindows,WindWindows,TSWindows,SnowWindows FROM TafStationFragments", conn).ExecuteReader();
				while (r.Read())
				{
					if (r.IsDBNull(0)) continue;
					TafFragments f;
					f.VisCeil = r.IsDBNull(1) ? "" : r.GetString(1);
					f.Wind    = r.IsDBNull(2) ? "" : r.GetString(2);
					f.TS      = r.IsDBNull(3) ? "" : r.GetString(3);
					f.Snow    = r.IsDBNull(4) ? "" : r.GetString(4);
					f.VisCeilWindows = DeserializeWindows(r.IsDBNull(5) ? "" : r.GetString(5));
					f.WindWindows    = DeserializeWindows(r.IsDBNull(6) ? "" : r.GetString(6));
					f.TSWindows      = DeserializeWindows(r.IsDBNull(7) ? "" : r.GetString(7));
					f.SnowWindows    = DeserializeWindows(r.IsDBNull(8) ? "" : r.GetString(8));
					result[r.GetString(0)] = f;
				}
				conn.Close();
			}
			catch { }
			return result;
		}

		private static bool ContainsRed(string fragment)
		{
			return !string.IsNullOrEmpty(fragment) && fragment.IndexOf("color:red", StringComparison.OrdinalIgnoreCase) >= 0;
		}

		// "yyyyMMddHHmm-yyyyMMddHHmm|yyyyMMddHHmm-yyyyMMddHHmm|..." — compact, sortable-free-form
		// text encoding for a List<TafWindow>, since TafStationFragments has no room for a proper
		// child table just for this.
		private const string WindowFormat = "yyyyMMddHHmm";

		private static string SerializeWindows(List<TafWindow> windows)
		{
			if (windows == null || windows.Count == 0) return "";
			List<string> parts = new List<string>();
			foreach (TafWindow w in windows)
				parts.Add(w.Start.ToString(WindowFormat, CultureInfo.InvariantCulture) + "-" + w.End.ToString(WindowFormat, CultureInfo.InvariantCulture));
			return string.Join("|", parts.ToArray());
		}

		private static List<TafWindow> DeserializeWindows(string serialized)
		{
			List<TafWindow> result = new List<TafWindow>();
			if (string.IsNullOrEmpty(serialized)) return result;
			foreach (string part in serialized.Split('|'))
			{
				string[] bounds = part.Split('-');
				if (bounds.Length != 2) continue;
				DateTime s, e;
				if (!DateTime.TryParseExact(bounds[0], WindowFormat, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out s)) continue;
				if (!DateTime.TryParseExact(bounds[1], WindowFormat, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out e)) continue;
				TafWindow w;
				w.Start = s;
				w.End = e;
				result.Add(w);
			}
			return result;
		}

		// ── Check vs Schedule: cross-references every CAT I (red)-flagged station against
		// today's Flight Schedule, same visual language as the Conflict/Email tab's own report
		// (airport header/diagram, flight chips) — reuses BuildAirportHeaderHtml/LoadFsFlights/
		// ConflictMatch/FsFlight directly (all defined in MainForm.Conflict.cs, same partial
		// class, so no visibility issue). Grouped into the same 4 categories as the full report
		// below (Ceiling/Vis, Wind, Thunderstorms, Snow) — a station's OWN flagged fragment for
		// that one category (frag.VisCeil/Wind/TS/Snow, already built by AnalyzeTaf) is shown
		// verbatim, so the "incriminated" text here is identical to the detailed section further
		// down the report, not a re-derived summary.
		//
		// Deliberately regenerated fresh on every call (not stored in TafReport.Html) — unlike
		// the dispatcher-editable ops-type body, this section embeds runway-diagram images, which
		// need raster+file:// for local viewing (rasterDiagrams=false/VML for the live tab, same
		// as Conflict's own live view) and raster+cid: for the emailed copy (Outlook's Word
		// engine only renders cid:-attached images in an HTML body). A single stored blob can't
		// serve both, so LoadTafReportIntoBrowser and Btn_sendTafReportClick each call this with
		// their own mode and prepend the result, wrapped in HTML comment markers
		// (CvsStartMarker/CvsEndMarker) so SaveEditedTafReport can strip it back out before
		// persisting a dispatcher's edit.
		private string BuildTafCheckVsScheduleHtml(bool rasterDiagrams, bool cidImages, out Dictionary<string, string> inlineImages)
		{
			inlineImages = new Dictionary<string, string>();

			Dictionary<string, TafFragments> fragByIcao = LoadTafStationFragments();
			List<FsFlight> flights = LoadFsFlights();

			// Only the next 24h counts — a flight scheduled tomorrow evening isn't meaningfully
			// "checked vs" a TAF issued for today.
			DateTime nowUtc = DateTime.UtcNow;
			DateTime next24hEnd = nowUtc.AddHours(24);

			string[] catLabels = { "Ceiling/Vis", "Wind", "Thunderstorms", "Snow" };
			string[] catBannerColors = { "#1565c0", "#00838f", "#c62828", "#5e35b1" };
			string[] catBadgeBg      = { "#bbdefb", "#b2ebf2", "#ffcdd2", "#e1bee7" };
			string[] catBadgeFg      = { "#0d47a1", "#00363a", "#7f0000", "#311b52" };

			// Wording matches the NOTAM Conflict tab's own introLine (MainForm.Conflict.cs) —
			// same 3-line structure, no "CAT I" jargon.
			StringBuilder sb = new StringBuilder();
			sb.Append("<p class=\"introLine\">Below is a summary of significant weather conditions in the coming 24h, cross-referenced against the flight schedule.<br>" +
				"For Origin/Dest, a match is checked at STD/STA &plusmn;1Hr, verified against the flagged TAF block's own hours.<br>" +
				"For ALTN use, the check is done at STD + flight time + diversion time &plusmn;2Hrs, verified against the flagged TAF block's own hours.</p>");
			if (flights.Count == 0)
				sb.Append("<div class=\"warnBanner\"><span class=\"warnIcon\">&#9888;</span>" +
					"Flight Schedule is empty — no flights to cross-reference. Load the Flight Schedule tab first.</div>");

			// Every category is always shown — same "the section stays visible, dimmed/Nil when
			// empty" idiom the NOTAM Report tab's Section()/AppendImpactRow already uses, rather
			// than silently disappearing when there's currently no impact (which reads as broken,
			// not as "confirmed nothing to report").
			for (int ci = 0; ci < 4; ci++)
			{
				StringBuilder catCards = new StringBuilder();
				int count = 0;
				if (flights.Count > 0)
				{
					foreach (KeyValuePair<string, TafFragments> kv in fragByIcao)
					{
						string icao = kv.Key;
						TafFragments frag = kv.Value;
						string fragText = ci == 0 ? frag.VisCeil : ci == 1 ? frag.Wind : ci == 2 ? frag.TS : frag.Snow;
						if (!ContainsRed(fragText)) continue;
						List<TafWindow> windows = ci == 0 ? frag.VisCeilWindows : ci == 1 ? frag.WindWindows : ci == 2 ? frag.TSWindows : frag.SnowWindows;

						// Only a flight whose relevant time falls both in the next 24h AND inside
						// the flagged block's own hours counts as a real impact — not just "any
						// flight today at that station".
						List<ConflictMatch> matches = BuildTafFlightMatches(icao, flights, windows, nowUtc, next24hEnd);
						if (matches.Count == 0) continue;

						string flightChips = "";
						foreach (ConflictMatch m in matches)
							flightChips += "<span class=\"flightChip flightChipAlert\">" + m.Text + "</span>";

						// Card border matches this category's own banner colour (catBannerColors[ci])
						// rather than the fixed red ".cardAlert" the Conflict report uses — Ceiling/Vis
						// cards are blue-bordered, Wind teal, Thunderstorms red, Snow purple.
						catCards.Append("<div class=\"card\" style=\"border:2px solid ").Append(catBannerColors[ci])
							.Append(";box-shadow:0 0 0 1px ").Append(catBannerColors[ci]).Append("\">");
						catCards.Append(BuildAirportHeaderHtml(icao, rasterDiagrams, cidImages, inlineImages));
						catCards.Append("<div class=\"body\">");
						catCards.Append(flightChips);
						catCards.Append("<div class=\"notamtext\">").Append(fragText).Append("</div>");
						catCards.Append("</div></div>");
						count++;
					}
				}

				// Banner per category, same visual language as the Conflict report's MISC/Not-ALTN
				// "X active in the coming 7 days" banners — a plain label + a solid count badge on
				// the right, dimmed to 0 (still shown) when nothing currently qualifies.
				sb.Append("<div class=\"banner\" style=\"border-left:4px solid ").Append(catBannerColors[ci]).Append("\">")
					.Append(catLabels[ci])
					.Append("<span class=\"bCount\" style=\"background:").Append(catBadgeBg[ci]).Append(";color:").Append(catBadgeFg[ci]).Append("\">")
					.Append(count).Append(count == 1 ? " station" : " stations").Append("</span></div>");

				if (count == 0)
					sb.Append("<div class=\"nilRow\">Nil — no time-verified impact in the next 24h.</div>");
				else
					sb.Append(catCards);
			}

			return CvsStartMarker + "<div id=\"tafCheckVsSchedule\" contenteditable=\"false\">" + sb + "</div>" + CvsEndMarker;
		}

		// Matches a station (by ICAO) against the Flight Schedule the same way the Conflict tab
		// does — Origin/Dest by IATA, plus filed Alt1/Alt2 diversion-estimate arrivals by ICAO —
		// but restricted to (a) the next 24h (nowUtc..next24hEnd) and (b) the flagged block's own
		// hours: a flight only counts as a real impact if its relevant time actually falls inside
		// at least one of the category's red TafWindows (MatchesWindows falls back to "any time
		// in the 24h range counts" if no window could be parsed at all, so an unparseable TAF
		// format never silently hides a real breach). Chip text format matches
		// Build_Conflict_Report()'s own ConflictMatch.Text exactly (MainForm.Conflict.cs).
		private List<ConflictMatch> BuildTafFlightMatches(string icao, List<FsFlight> flights, List<TafWindow> windows, DateTime nowUtc, DateTime next24hEnd)
		{
			List<ConflictMatch> matches = new List<ConflictMatch>();
			string iata = GetIATA(icao);
			foreach (FsFlight f in flights)
			{
				if (iata != "" && f.Origin == iata && f.HasStd && InNext24h(f.Std, nowUtc, next24hEnd) && MatchesWindows(f.Std, windows))
					matches.Add(new ConflictMatch { FltlegId = f.FltlegID,
						Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — origin — STD " + FormatUtc(f.Std) + "Z" });
				if (iata != "" && f.Dest == iata && f.HasSta && InNext24h(f.Sta, nowUtc, next24hEnd) && MatchesWindows(f.Sta, windows))
					matches.Add(new ConflictMatch { FltlegId = f.FltlegID,
						Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — destination — STA " + FormatUtc(f.Sta) + "Z" });
				if (f.Alt1 == icao && f.FlightTimeMin > 0 && f.Alt1TimeMin > 0)
				{
					DateTime altArrival = f.Std.AddMinutes(f.FlightTimeMin + f.Alt1TimeMin);
					if (InNext24h(altArrival, nowUtc, next24hEnd) && MatchesWindows(altArrival, windows))
						matches.Add(new ConflictMatch { FltlegId = f.FltlegID,
							Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — alternate 1 — est. diversion arrival " + FormatUtc(altArrival) + "Z" });
				}
				if (f.Alt2 == icao && f.FlightTimeMin > 0 && f.Alt2TimeMin > 0)
				{
					DateTime altArrival = f.Std.AddMinutes(f.FlightTimeMin + f.Alt2TimeMin);
					if (InNext24h(altArrival, nowUtc, next24hEnd) && MatchesWindows(altArrival, windows))
						matches.Add(new ConflictMatch { FltlegId = f.FltlegID,
							Text = f.Callsign + " " + f.Origin + "-" + f.Dest + " — alternate 2 — est. diversion arrival " + FormatUtc(altArrival) + "Z" });
				}
			}
			return matches;
		}

		private static bool InNext24h(DateTime t, DateTime nowUtc, DateTime next24hEnd) { return t >= nowUtc && t <= next24hEnd; }

		private static bool MatchesWindows(DateTime t, List<TafWindow> windows)
		{
			if (windows == null || windows.Count == 0) return true; // nothing parsed -> don't hide a real breach
			foreach (TafWindow w in windows)
				if (t >= w.Start && t <= w.End) return true;
			return false;
		}

		// Subset of MainForm.Conflict.cs's own <style> block (same "twin, not shared" precedent
		// used throughout this codebase for per-report HTML, e.g. BuildRwySvgMini) — just the
		// classes BuildAirportHeaderHtml/the card/chip/banner markup above actually reference.
		// The TAF report has its own <html><head><style> wrapper (LoadTafReportIntoBrowser/
		// Btn_sendTafReportClick), unlike the Conflict report which never needed this section.
		private static string TafCheckVsScheduleCss()
		{
			return
				"v\\:*{behavior:url(#default#VML)}" +
				".introLine{font-size:15.5px;font-weight:600;color:#0d47a1;background:#eef3fb;border-left:3px solid #1565c0;padding:8px 12px;border-radius:0 4px 4px 0;margin:0 0 16px 0;line-height:1.4}" +
				// Banner + solid count badge — same shape as the Conflict report's MISC/Not-ALTN
				// "X active in the coming 7 days" banners (MainForm.Conflict.cs), reused here so
				// the 4 Check vs Schedule categories read as the same visual family.
				".banner{position:relative;display:block;padding:10px 14px;border-radius:6px;margin:16px 0 8px 0;background:#eceff1;font-size:14px;font-weight:600;color:#37474f}" +
				".banner .bCount{position:absolute;top:8px;right:10px;text-align:center;font-size:13px;font-weight:bold;padding:4px 10px;border-radius:3px}" +
				".nilRow{color:#78909c;font-size:13px;font-style:italic;margin:0 0 12px 4px}" +
				".card{border:1px solid #cfd8dc;border-radius:8px;overflow:hidden;margin:0 0 18px 0}" +
				".ahead{background:#263238;padding:14px 18px;position:relative}" +
				".icao{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".sub{font-size:13px;color:#78909c;margin-top:2px}" +
				".apname{font-size:13px;color:#90a4ae;margin-top:1px}" +
				".blk{font-size:12px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:6px 12px;margin-right:8px;vertical-align:top}" +
				".rwytable{margin-top:8px}" +
				".rwyline{white-space:nowrap;line-height:1.7}" +
				".diagram{position:absolute;top:10px;right:60px}" +
				".aheadTable td{vertical-align:top}" +
				".body{padding:12px 18px}" +
				".flightChip{display:inline-block;background:#fbe9e7;color:#4e342e;font-size:12px;padding:5px 10px;border-radius:6px;margin:0 8px 8px 0}" +
				".flightChipAlert{background:#c62828;color:#fff;font-weight:bold}" +
				".notamtext{background:#f5f5f5;border-radius:6px;padding:10px 12px;font-family:'Courier New',monospace;font-size:12.5px;white-space:pre-wrap;line-height:1.6}" +
				".warnBanner{background:#fff3e0;color:#7a4a00;border:1px solid #ffcc80;border-radius:6px;padding:10px 14px;margin:0 0 16px 0;font-size:13px}" +
				".warnIcon{margin-right:8px}";
		}

		// ── Tab lifecycle ──

		void TafAnalysisTabEnter(object sender, EventArgs e)
		{
			if (_tafTopBar == null) BuildTafTopBar();
			// Check vs Schedule depends on FlightSchedule — same throttled (5 min), silent-for-
			// Readers trigger the Flight Schedule and Conflict tabs already use
			// (TryAutoRefreshFlightSchedule, MainForm.FlightSchedule.cs). The refresh itself runs
			// on a BackgroundWorker, so this render pass still builds from whatever's already
			// stored; RefreshFlightSchedule()'s own completion handler re-renders this tab if it's
			// still the active one once the fresh data actually lands.
			TryAutoRefreshFlightSchedule();
			LoadTafReportIntoBrowser();
		}

		void TafAnalysisTabLeave(object sender, EventArgs e)
		{
			SaveEditedTafReport();
		}

		private void LoadTafReportIntoBrowser()
		{
			string timeIssued;
			string storedBody = LoadTafReportHtml(out timeIssued);
			if (storedBody == "")
				storedBody = "No TAF analysis yet — click \"TAF Analysis\" to download and analyze.";

			// Check vs Schedule is never part of the stored/editable body — it's rebuilt fresh
			// every render (see BuildTafCheckVsScheduleHtml's header comment) and prepended here.
			// Live tab uses VML (rasterDiagrams=false), same as the Conflict tab's own live view.
			Dictionary<string, string> unusedInlineImages;
			string checkHtml = BuildTafCheckVsScheduleHtml(false, false, out unusedInlineImages);

			// xmlns:v is required on the root <html> tag for the VML runway diagrams
			// (BuildRwySvg/BuildRwySvgGeo) to render at all — without it MSHTML silently drops
			// every <v:*> shape while still rendering the plain-<div> QFU labels next to them,
			// which reads as "the diagram is half-broken" (QFU shown, no runway lines) rather
			// than an obvious failure. Same wrapper Conflict/Email's own live tab already uses
			// (MainForm.Conflict.cs, BuildConflictReportHtml).
			string html = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\"><head><style>" + TafCheckVsScheduleCss() + "</style></head>" +
				"<body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:13px\">" +
				checkHtml + storedBody + "</body></html>";

			_tafBodyLoaded = false;
			Web_TafAnalysis.DocumentCompleted -= Web_TafAnalysisDocumentCompleted;
			Web_TafAnalysis.DocumentCompleted += Web_TafAnalysisDocumentCompleted;
			Web_TafAnalysis.DocumentText = html;
		}

		// contentEditable can only be applied once MSHTML has actually finished loading the
		// document — setting it right after DocumentText= is a no-op. Only a Writer's browser
		// is made editable; a Reader's stays genuinely read-only rather than looking editable
		// and silently discarding edits.
		void Web_TafAnalysisDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
		{
			if (_tafBodyLoaded) return;
			_tafBodyLoaded = true;
			try
			{
				if (IsWriter && Web_TafAnalysis.Document != null && Web_TafAnalysis.Document.Body != null)
					Web_TafAnalysis.Document.Body.SetAttribute("contentEditable", "true");
			}
			catch { }
		}

		// Persists whatever the dispatcher currently sees/edited back to the shared TafReport
		// row. Called on tab Leave and again right before Send. No-ops for a Reader (whose
		// WebBrowser was never made contentEditable) or before anything has ever loaded.
		private void SaveEditedTafReport()
		{
			if (!IsWriter) return;
			try
			{
				if (Web_TafAnalysis.Document == null || Web_TafAnalysis.Document.Body == null) return;
				string html = Web_TafAnalysis.Document.Body.InnerHtml;
				if (html == null) return;
				html = StripCheckVsScheduleBlock(html);
				string timeIssued;
				LoadTafReportHtml(out timeIssued); // keep the existing "last analysis" timestamp — editing text isn't a new analysis
				SaveTafReportHtml(html, timeIssued);
			}
			catch { }
		}

		// The Check vs Schedule block LoadTafReportIntoBrowser() prepends is marked with plain
		// HTML comments (not a matched <div id="..."> pair — MSHTML's InnerHtml has no reliable
		// "find the matching closing tag" primitive via regex once the block itself contains
		// nested <div>s, as every airport card here does) so it can be sliced back out with a
		// simple substring op before a dispatcher's edit gets persisted — otherwise every Leave
		// would re-save a stale copy of this auto-generated section into TafReport.Html.
		private const string CvsStartMarker = "<!--CVS_START-->";
		private const string CvsEndMarker = "<!--CVS_END-->";

		private static string StripCheckVsScheduleBlock(string html)
		{
			int s = html.IndexOf(CvsStartMarker, StringComparison.Ordinal);
			int e = html.IndexOf(CvsEndMarker, StringComparison.Ordinal);
			if (s < 0 || e < 0 || e < s) return html;
			return html.Substring(0, s) + html.Substring(e + CvsEndMarker.Length);
		}

		// ── Top bar: TAF Analysis / Send / Recipients / Attach Image + threshold fields ──

		private void BuildTafTopBar()
		{
			_tafTopBar = new Panel { Dock = DockStyle.Top, Height = 130 };
			// Parented immediately, before any child controls are added — Dock=Top only takes
			// on the tab's real width once the panel actually has a parent, and the right-aligned
			// buttons below need that real width (via Anchor) at the moment they're added, not
			// the Panel's ~200px design-time default. Adding children first and parenting the
			// panel last (the original order) anchored them against that bogus 200px width
			// instead, pushing them off the visible area entirely.
			tabPage_TafAnalysis.Controls.Add(_tafTopBar);

			Button analyzeBtn = new Button
			{
				Top = 8, Left = 14, Width = 140, Height = 30, Text = "TAF Analysis",
				Font = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold),
				BackColor = Color.SeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			analyzeBtn.Click += delegate { DownloadAndAnalyzeTafs(); };
			_tafTopBar.Controls.Add(analyzeBtn);

			// Send / Update Recipients — right-aligned, anchored so they stay pinned to the
			// right edge if the tab is ever resized. ClientSize.Width is already valid here since
			// this is built lazily on the tab's first Enter, after the form has been shown/sized.
			int rightEdge = tabPage_TafAnalysis.ClientSize.Width;
			int recipientsWidth = 170, sendWidth = 160, rightMargin = 14, gap = 10;

			Button recipientsBtn = new Button
			{
				Top = 8, Left = rightEdge - rightMargin - recipientsWidth, Width = recipientsWidth, Height = 30,
				Text = "Update TAF Recipients", Anchor = AnchorStyles.Top | AnchorStyles.Right,
				BackColor = Color.SkyBlue, ForeColor = Color.Black, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			recipientsBtn.Click += delegate { ShowTafRecipientsDialog(); };
			_tafTopBar.Controls.Add(recipientsBtn);

			Button sendBtn = new Button
			{
				Top = 8, Left = recipientsBtn.Left - gap - sendWidth, Width = sendWidth, Height = 30,
				Text = "✉  Send TAF Report", Anchor = AnchorStyles.Top | AnchorStyles.Right,
				BackColor = Color.FromArgb(21, 101, 192), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			sendBtn.Click += Btn_sendTafReportClick;
			_tafTopBar.Controls.Add(sendBtn);

			int labelTop = 46, boxTop = 62, fieldWidth = 55, spacing = 92, x = 14;
			Font smallFont = new Font("Microsoft Sans Serif", 7.5f);

			Label lVisCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Vis CAT I (m)" };
			_tafTopBar.Controls.Add(lVisCatI);
			_tafVisCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafVisCatIm.ToString() };
			_tafTopBar.Controls.Add(_tafVisCatIBox);
			x += spacing;

			Label lVisAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Vis Adv (m)" };
			_tafTopBar.Controls.Add(lVisAdv);
			_tafVisAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafVisAdvisoryM.ToString() };
			_tafTopBar.Controls.Add(_tafVisAdvBox);
			x += spacing;

			Label lCeilCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Ceil CAT I (x100ft)" };
			_tafTopBar.Controls.Add(lCeilCatI);
			_tafCeilCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafCeilCatIHundredFt.ToString() };
			_tafTopBar.Controls.Add(_tafCeilCatIBox);
			x += spacing;

			Label lCeilAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Ceil Adv (x100ft)" };
			_tafTopBar.Controls.Add(lCeilAdv);
			_tafCeilAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafCeilAdvisoryHundredFt.ToString() };
			_tafTopBar.Controls.Add(_tafCeilAdvBox);
			x += spacing;

			Label lWindKtCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind CAT I (kt)" };
			_tafTopBar.Controls.Add(lWindKtCatI);
			_tafWindKtCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindCatIKt.ToString() };
			_tafTopBar.Controls.Add(_tafWindKtCatIBox);
			x += spacing;

			Label lWindKtAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind Adv (kt)" };
			_tafTopBar.Controls.Add(lWindKtAdv);
			_tafWindKtAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindAdvisoryKt.ToString() };
			_tafTopBar.Controls.Add(_tafWindKtAdvBox);
			x += spacing;

			Label lWindMpsCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind CAT I (mps)" };
			_tafTopBar.Controls.Add(lWindMpsCatI);
			_tafWindMpsCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindCatIMps.ToString() };
			_tafTopBar.Controls.Add(_tafWindMpsCatIBox);
			x += spacing;

			Label lWindMpsAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind Adv (mps)" };
			_tafTopBar.Controls.Add(lWindMpsAdv);
			_tafWindMpsAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindAdvisoryMps.ToString() };
			_tafTopBar.Controls.Add(_tafWindMpsAdvBox);
			x += spacing;

			Button saveThresholds = new Button
			{
				Top = 58, Left = x, Width = 120, Height = 26, Text = "Save Thresholds",
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			saveThresholds.Click += delegate { SaveTafThresholds(); };
			_tafTopBar.Controls.Add(saveThresholds);

			// Row 3, below the threshold fields: Attach Image, then a small text-editing toolbar
			// (Bold/Italic/Underline + 4 color swatches) that applies to whatever's currently
			// selected in the editable report body via the WebBrowser's own Document.ExecCommand
			// — the standard MSHTML rich-text-toolbar technique, no JS bridge needed.
			int row3Top = 94;

			Button attachBtn = new Button
			{
				Top = row3Top, Left = 14, Width = 130, Height = 28, Text = "Attach Image",
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			attachBtn.Click += delegate { AttachImageToTafReport(); };
			_tafTopBar.Controls.Add(attachBtn);

			int fx = 156;
			Font boldFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold);
			Font italicFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Italic);
			Font underlineFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Underline);

			Button boldBtn = new Button { Top = row3Top, Left = fx, Width = 30, Height = 28, Text = "B", Font = boldFont,
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false };
			boldBtn.Click += delegate { ExecTafFormat("Bold", null); };
			_tafTopBar.Controls.Add(boldBtn);
			fx += 34;

			Button italicBtn = new Button { Top = row3Top, Left = fx, Width = 30, Height = 28, Text = "I", Font = italicFont,
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false };
			italicBtn.Click += delegate { ExecTafFormat("Italic", null); };
			_tafTopBar.Controls.Add(italicBtn);
			fx += 34;

			Button underlineBtn = new Button { Top = row3Top, Left = fx, Width = 30, Height = 28, Text = "U", Font = underlineFont,
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false };
			underlineBtn.Click += delegate { ExecTafFormat("Underline", null); };
			_tafTopBar.Controls.Add(underlineBtn);
			fx += 44;

			fx = AddTafColorSwatch(fx, row3Top, Color.Black, "black");
			fx = AddTafColorSwatch(fx, row3Top, Color.Red, "red");
			fx = AddTafColorSwatch(fx, row3Top, Color.FromArgb(0xB8, 0x86, 0x0B), "#B8860B");
			fx = AddTafColorSwatch(fx, row3Top, Color.Blue, "blue");

			Web_TafAnalysis.BringToFront();
		}

		// One small square color-swatch button applying ForeColor to the current selection in
		// the editable report. Returns the Left for the next swatch.
		private int AddTafColorSwatch(int left, int top, Color color, string execColor)
		{
			Button swatch = new Button { Top = top, Left = left, Width = 24, Height = 28, Text = "",
				BackColor = color, FlatStyle = FlatStyle.Flat };
			swatch.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
			swatch.Click += delegate { ExecTafFormat("ForeColor", execColor); };
			_tafTopBar.Controls.Add(swatch);
			return left + 30;
		}

		// Applies a rich-text command to whatever's currently selected in the editable report
		// body (Bold/Italic/Underline/ForeColor), then persists the result — same
		// EnsureWriterOrWarn() gate as every other edit to the shared report.
		private void ExecTafFormat(string command, string value)
		{
			if (!EnsureWriterOrWarn()) return;
			try
			{
				if (Web_TafAnalysis.Document == null) return;
				Web_TafAnalysis.Document.ExecCommand(command, false, value);
				SaveEditedTafReport();
			}
			catch { }
		}

		private void SaveTafThresholds()
		{
			if (!EnsureWriterOrWarn()) return;
			int v;
			if (int.TryParse(_tafVisCatIBox.Text.Trim(), out v) && v > 0) _tafVisCatIm = v;
			if (int.TryParse(_tafVisAdvBox.Text.Trim(), out v) && v > 0) _tafVisAdvisoryM = v;
			if (int.TryParse(_tafCeilCatIBox.Text.Trim(), out v) && v > 0) _tafCeilCatIHundredFt = v;
			if (int.TryParse(_tafCeilAdvBox.Text.Trim(), out v) && v > 0) _tafCeilAdvisoryHundredFt = v;
			if (int.TryParse(_tafWindKtCatIBox.Text.Trim(), out v) && v > 0) _tafWindCatIKt = v;
			if (int.TryParse(_tafWindKtAdvBox.Text.Trim(), out v) && v > 0) _tafWindAdvisoryKt = v;
			if (int.TryParse(_tafWindMpsCatIBox.Text.Trim(), out v) && v > 0) _tafWindCatIMps = v;
			if (int.TryParse(_tafWindMpsAdvBox.Text.Trim(), out v) && v > 0) _tafWindAdvisoryMps = v;
			SaveArchiveConfig();
			MessageBox.Show("Thresholds saved.", "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		// ── Download + analyze ──

		public void DownloadAndAnalyzeTafs()
		{
			if (!EnsureWriterOrWarn()) return;
			if (_stationsCache == null) LoadStationsCache();

			List<string> icaos = new List<string>();
			foreach (var entry in _stationsCache)
			{
				string[] row = entry.Value;
				if (row[1] == "Yes" || row[2] == "Yes" || row[3] == "Yes")
					icaos.Add(entry.Key);
			}
			if (icaos.Count == 0)
			{
				MessageBox.Show("No airports flagged LH/FedEx/Charters on the Airport List.", "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Same ddMMMyy computation GetXML() (MainForm.NotamData.cs) already uses.
			string todayStr = DateTime.Now.ToString("yyyyMMdd");
			string stringToday = todayStr.Substring(6, 2) + MonthAbbrev(todayStr.Substring(4, 2)) + todayStr.Substring(2, 2);

			string request = string.Join("-", icaos.ToArray());
			string url = _metBaseUrl.TrimEnd('/') + "/?METHOD=getAdHocMET&METTYPE=FT&REQUEST=" + request +
				"&PERIODSTART=" + stringToday + "&PERIODEND=" + stringToday;

			string xml;
			try
			{
				using (WebClient wc = new WebClient())
					xml = wc.DownloadString(url);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to download TAFs:\n" + ex.Message, "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Parsed by local name only (same fix FetchBriefingDetails / MainForm.FlightSchedule.cs
			// already applies to this ARINC 633 namespace family) — matching on a fully-qualified
			// XName against the wrapping http://aeec.aviation-ia.net/633 namespace silently finds
			// nothing.
			Dictionary<string, string> rawByIcao = new Dictionary<string, string>();
			Dictionary<string, TafWindow> validByIcao = new Dictionary<string, TafWindow>();
			try
			{
				XDocument doc = XDocument.Parse(xml);
				foreach (XElement bulletin in doc.Descendants().Where(el => el.Name.LocalName == "WeatherBulletin"))
				{
					XElement icaoEl = bulletin.Descendants().FirstOrDefault(el => el.Name.LocalName == "AirportICAOCode");
					if (icaoEl == null || string.IsNullOrEmpty(icaoEl.Value)) continue;
					string icao = icaoEl.Value.Trim();

					StringBuilder sb = new StringBuilder();
					foreach (XElement textEl in bulletin.Descendants().Where(el => el.Name.LocalName == "Text"))
						sb.Append(textEl.Value).Append(" ");
					string raw = sb.ToString().Trim().TrimEnd('=').Trim();
					if (raw == "") continue;
					rawByIcao[icao] = raw; // last bulletin for a station wins if more than one is returned

					// forecastStartTime/forecastEndTime (e.g. "2026-08-13T00:00:00.000Z") give the
					// TAF's own overall validity window — needed by AnalyzeTaf to derive real UTC
					// times for the base/initial group and any FM (open-ended) trend group.
					TafWindow valid;
					valid.Start = DateTime.UtcNow.Date;
					valid.End = valid.Start.AddDays(2);
					XElement forecastEl = bulletin.Descendants().FirstOrDefault(el => el.Name.LocalName == "Forecast");
					if (forecastEl != null)
					{
						XAttribute fs = forecastEl.Attribute("forecastStartTime");
						XAttribute fe = forecastEl.Attribute("forecastEndTime");
						DateTime parsed;
						if (fs != null && DateTime.TryParse(fs.Value, CultureInfo.InvariantCulture,
							DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out parsed)) valid.Start = parsed;
						if (fe != null && DateTime.TryParse(fe.Value, CultureInfo.InvariantCulture,
							DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out parsed)) valid.End = parsed;
					}
					validByIcao[icao] = valid;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to parse TAF response:\n" + ex.Message, "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			Dictionary<string, TafFragments> fragByIcao = new Dictionary<string, TafFragments>(StringComparer.OrdinalIgnoreCase);
			foreach (string icao in icaos)
			{
				string raw;
				if (!rawByIcao.TryGetValue(icao, out raw)) continue; // no FT returned for this station
				TafWindow valid = validByIcao[icao];
				fragByIcao[icao] = AnalyzeTaf(raw, valid.Start, valid.End);
			}

			// Every station with at least one red (CAT I) fragment is snapshotted into
			// TafStationFragments — the only data BuildTafCheckVsScheduleHtml needs to rebuild
			// the "Check vs schedule" section on demand, both for the live tab and at Send time,
			// without having to re-download/re-parse the raw TAF text.
			SaveTafStationFragments(fragByIcao);

			// Organised like the legacy standalone app's db_read(): one section per ops type
			// (Long Haul / FedEx / Charters, same title colours), each broken into the same 4
			// sub-categories (Ceiling & Visibility, Wind, Thunderstorms, Snow) — a station only
			// appears under a category if that category actually flagged something for it.
			// This is the dispatcher-editable body only — no outer <html>/<body> wrapper, since
			// SaveEditedTafReport() likewise only ever stores Document.Body.InnerHtml; the
			// wrapper (and the auto-generated Check vs Schedule section) is added back on at
			// render/send time by LoadTafReportIntoBrowser/Btn_sendTafReportClick.
			StringBuilder report = new StringBuilder();
			// Same blue banner style as the Check vs Schedule / Conflict report's own introLine
			// (defined in TafCheckVsScheduleCss(), available here too since this body is always
			// concatenated into that same document) — reused as-is rather than a one-off style.
			report.Append("<p class=\"introLine\">NETWORK OVERVIEW</p>");
			report.Append("<span style=\"font-weight:bold;color:#888\">Last analysis: " +
				DateTime.Now.ToString("dd/MM/yyyy HH:mm") + " (local)</span><hr />");
			report.Append("<table style=\"text-align:left; font-size:13px\">");
			AppendOpsSection(report, "LH", "Long Haul Ops :", "RoyalBlue", icaos, fragByIcao);
			AppendOpsSection(report, "FedEx", "FedEx Ops :", "DarkMagenta", icaos, fragByIcao);
			AppendOpsSection(report, "Charters", "Charters Ops :", "Green", icaos, fragByIcao);
			report.Append("</table>");

			string timeIssued = DateTime.Now.ToString("dd/MM/yyyy HH:mm") + " (local)";
			SaveTafReportHtml(report.ToString(), timeIssued);
			_tafAnalysisRunThisSession = true;
			LoadTafReportIntoBrowser();
		}

		// Per-station threshold-analysis result, one HTML fragment per category — mirrors the
		// legacy app's Vis_Ceiling/Wind/TS/Snow accumulator columns (TAF_analysis table there).
		// Each fragment only ever contains the tokens that actually matched that category (same
		// as the legacy accumulators), colored red/amber/blue exactly like AnalyzeTaf's combined
		// mode did before this change — it's the same underlying per-token analysis, just routed
		// into 4 separate buckets instead of one combined colored line.
		private struct TafFragments
		{
			public string VisCeil, Wind, TS, Snow;
			// One entry per RED (CAT I) block actually appended to the matching fragment above —
			// the real-world UTC time window that block's trend group covers (e.g. "BECMG
			// 1310/1312" -> 13th 10:00-12:00Z), used by Check vs Schedule to verify a flight's
			// STD/STA/diversion time actually falls inside the flagged period, not just "today".
			public List<TafWindow> VisCeilWindows, WindWindows, TSWindows, SnowWindows;
		}

		// A UTC time window derived from one flagged TAF trend-group block.
		private struct TafWindow
		{
			public DateTime Start, End;
		}

		// One ops-type section (title + hr), each broken into the 4 sub-category blocks below.
		private static void AppendOpsSection(StringBuilder report, string opsType, string title, string color,
			List<string> icaos, Dictionary<string, TafFragments> fragByIcao)
		{
			report.Append("<tr><th colspan=\"5\"><span style=\"font-weight:bold; font-size:16px; color:" + color + ";\">" +
				title + "</span></th></tr>");
			AppendCategory(report, "Ceiling & Visibility", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.VisCeil; });
			AppendCategory(report, "Wind", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.Wind; });
			AppendCategory(report, "Thunderstorms", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.TS; });
			AppendCategory(report, "Snow", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.Snow; });
			report.Append("<tr><th colspan=\"5\"><hr /></th></tr>");
		}

		private static void AppendCategory(StringBuilder report, string categoryTitle, string opsType, List<string> icaos,
			Dictionary<string, TafFragments> fragByIcao, Func<TafFragments, string> selector)
		{
			report.Append("<tr><th colspan=\"5\"><b>" + categoryTitle + "</b></th></tr><tr><th colspan=\"5\">");
			report.Append("<span style=\"font-family:'Courier New'; font-weight:normal;\">");
			bool any = false;
			foreach (string icao in icaos)
			{
				if (IsOpsType(opsType, icao) != "Yes") continue;
				TafFragments frag;
				if (!fragByIcao.TryGetValue(icao, out frag)) continue;
				string val = selector(frag);
				if (string.IsNullOrEmpty(val)) continue;
				string iata = GetIATA(icao);
				report.Append("<b style=\"color:DarkBlue;\">" + icao + (string.IsNullOrEmpty(iata) ? "" : " - " + iata) +
					" : </b>" + val + "<br />");
				any = true;
			}
			if (!any) report.Append("Nil");
			report.Append("</span></th></tr>");
		}

		// Direct port of the legacy "TAF analysis" app's per-station algorithm
		// (TAF analysis/Dispatch/MainForm.cs, Btn_addTAFClick, lines ~180-407) — the
		// station/time splitting that code needs is unnecessary here since the webservice XML
		// already hands back one clean forecast string per station. Like the legacy app, only
		// the specific value(s) that breached a threshold are shown per category (not the whole
		// TAF clause/token) — e.g. "OVC002" or "34012KT", colored by tier — routed into the same
		// 4 category buckets (VisCeil/Wind/TS/Snow) the report groups by.
		private static TafFragments AnalyzeTaf(string raw, DateTime validFrom, DateTime validTo)
		{
			// Display granularity matches the legacy app exactly: each flagged trend group is
			// shown as its WHOLE block/clause (e.g. "BECMG 1310/1312 34012KT", not just
			// "34012KT"), so the hourly period and rest of that block stay attached to whatever
			// triggered it. Blocks are appended to their category bucket in the TAF's own
			// chronological order, so a breach followed later by its recovery (BECMG/FM back
			// above threshold) both land in the same fragment, one after another — the
			// surrounding trend is visible without extra bookkeeping. Colors match the legacy
			// app: red for a CAT I breach, blue for a trend-group recovery — plus amber
			// (`#B8860B`) for the advisory tier, so a below/above-threshold-but-not-CAT-I block
			// is still visually distinguishable from unflagged text. TS/Snow are always red
			// (unconditional flags — any thunderstorm/snow/freezing-precip keyword is a hazard,
			// no threshold tiers involved).
			// Trend markers: BECMG/TEMPO/PROB30/PROB40/FM groups. The pattern's capturing group
			// means Regex.Split also returns each matched marker as its own array element.
			string patternTrend = @"(BECMG [0-9]{4}|TEMPO [0-9]{4}|PROB30 [0-9]{4}|PROB40 [0-9]{4}|FM[0-9]{4})";
			string[] parts = Regex.Split(raw, patternTrend);

			// PROB30/PROB40 + TEMPO reattachment: for "PROB30 TEMPO 1315/1318 ...", the TEMPO
			// alternative matches starting at "TEMPO", leaving "PROB30 " dangling on the END of
			// the PRECEDING chunk instead of prefixed onto the delimiter — ported verbatim from
			// the legacy app's two-pass fixup (lines 194-218 there).
			for (int i = 0; i < parts.Length - 1; i++)
			{
				int len = parts[i].Length;
				string end7 = len > 7 ? parts[i].Substring(len - 7, 7) : "";
				if (end7 == "PROB30 " || end7 == "PROB40 ")
				{
					parts[i] = parts[i].Substring(0, len - 7);
					parts[i + 1] = end7 + parts[i + 1];
					continue;
				}
				string end10 = len > 10 ? parts[i].Substring(len - 10, 7) : "";
				if (end10 == "PROB30 " || end10 == "PROB40 ")
				{
					parts[i] = parts[i].Substring(0, len - 10);
					parts[i + 1] = end10 + parts[i + 1];
				}
			}

			// Break into display tokens — each trend group starts its own token unless it's a
			// bare continuation (leading "/" or "0"), matching the legacy app's <br/> heuristic.
			StringBuilder joined = new StringBuilder();
			foreach (string part in parts)
			{
				string lead = part.Length > 0 ? part.Substring(0, 1) : "-";
				if (lead == "/" || lead == "0") joined.Append(part);
				else joined.Append("<br />").Append(part);
			}
			string[] tokens = Regex.Split(joined.ToString(), "<br />");

			// Per-token bookkeeping needed to compute windows AFTER the fact (second pass below) —
			// a token's own text only ever gives its OWN transition/closed bounds; the real
			// question ("until when do these conditions actually apply?") for a BECMG/FM/base
			// group depends on what the NEXT persisting group says, which isn't known until the
			// whole TAF has been walked once.
			List<TafTokenMeta> metas = new List<TafTokenMeta>();

			// Each colored contribution is tagged with the index of the token it came from,
			// instead of being appended straight into a StringBuilder — display-time filtering
			// (below) needs to be able to drop a token's contribution from every category after
			// the fact, once every token's own bounds are known.
			List<FragEntry> visCeilEntries = new List<FragEntry>(), windEntries = new List<FragEntry>(),
				tsEntries = new List<FragEntry>(), snowEntries = new List<FragEntry>();

			bool visTrend = false, windTrend = false;
			int tokenIdx = 0;
			foreach (string token in tokens)
			{
				if (token.Trim() == "") continue;
				string lead2 = token.Length >= 2 ? token.Substring(0, 2) : "";
				// "Persisting": BECMG/FM/the base(initial) group — conditions it describes remain
				// in effect until superseded by the NEXT persisting group. TEMPO/PROB30/PROB40 are
				// NOT persisting — they're a closed, temporary window; once it ends, conditions
				// revert to whatever the surrounding persisting group already said.
				bool isPersisting = lead2 == "BE" || lead2 == "FM" || (token.Length > 0 && char.IsDigit(token[0]));

				TafTokenMeta meta = new TafTokenMeta();
				meta.IsPersisting = isPersisting;
				ParseOwnBounds(token, validFrom, validTo, out meta.OwnStart, out meta.OwnEnd);

				// Vis + ceiling flags are aggregated across every match in this token, then the
				// WHOLE block is appended once (red wins over the advisory tier over a trend
				// recovery) — exactly the legacy app's ceilCatI/ceilTresh/visCatI/visTresh
				// aggregate-then-decide-once pattern (lines 300-317 there), not a per-match append.
				bool visCatI = false, visTresh = false, ceilCatI = false, ceilTresh = false;
				foreach (Match m in Regex.Matches(token, @"(?<= )\b[0-9]{4,4}\b(?= )"))
				{
					int val = int.Parse(m.Value);
					if (val < _tafVisCatIm) { visCatI = true; if (isPersisting) visTrend = true; }
					else if (val <= _tafVisAdvisoryM) { visTresh = true; if (isPersisting) visTrend = true; }
				}
				foreach (Match m in Regex.Matches(token, @"(?<=BKN|OVC|VV)[0-9]{3}"))
				{
					int val = int.Parse(m.Value);
					if (val <= _tafCeilCatIHundredFt) { ceilCatI = true; if (isPersisting) visTrend = true; }
					else if (val <= _tafCeilAdvisoryHundredFt) { ceilTresh = true; if (isPersisting) visTrend = true; }
				}
				if (ceilCatI || visCatI) { visCeilEntries.Add(new FragEntry(tokenIdx, BlockHtml(token, "red"))); meta.VisCeilRed = true; }
				else if (ceilTresh || visTresh) visCeilEntries.Add(new FragEntry(tokenIdx, BlockHtml(token, "#B8860B")));
				else if (isPersisting && visTrend) { visCeilEntries.Add(new FragEntry(tokenIdx, BlockHtml(token, "blue"))); visTrend = false; }

				// Wind KT/MPS append per match (as the legacy app does), since a token normally
				// carries at most one wind group. The speed/gust portion of the matched wind
				// group (everything after the 3-char direction, e.g. "35KT" or "25G35KT") is
				// bolded within the block — the direction itself stays plain weight.
				foreach (Match m in Regex.Matches(token, @"([a-zA-Z0-9]{5,8})KT"))
				{
					string w = m.Value;
					string spd = w.Length == 10 ? w.Substring(6, 2) : w.Substring(3, 2);
					int kt;
					if (int.TryParse(spd, out kt))
					{
						string boldedToken = BoldWindSpeed(token, m);
						if (kt > _tafWindCatIKt) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "red"))); meta.WindRed = true; if (isPersisting) windTrend = true; }
						else if (kt >= _tafWindAdvisoryKt) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "#B8860B"))); if (isPersisting) windTrend = true; }
						else if (isPersisting && windTrend) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "blue"))); windTrend = false; }
					}
				}
				foreach (Match m in Regex.Matches(token, @"([a-zA-Z0-9]{5,8})MPS"))
				{
					string w = m.Value;
					string spd = w.Length == 11 ? w.Substring(6, 2) : w.Substring(3, 2);
					int mps;
					if (int.TryParse(spd, out mps))
					{
						string boldedToken = BoldWindSpeed(token, m);
						if (mps > _tafWindCatIMps) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "red"))); meta.WindRed = true; if (isPersisting) windTrend = true; }
						else if (mps >= _tafWindAdvisoryMps) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "#B8860B"))); if (isPersisting) windTrend = true; }
						else if (isPersisting && windTrend) { windEntries.Add(new FragEntry(tokenIdx, BlockHtml(boldedToken, "blue"))); windTrend = false; }
					}
				}
				// TS / snow / freezing precip — always flagged red, no threshold tiers
				// (unconditional, same trigger set as the legacy app's TS+=valueBr / SN+=valueBr).
				// Highlight the WHOLE weather group (e.g. "TSRA", "TSGS", "VCTS", "+TSRA"), not
				// just the bare "TS"/"SN" substring — \w*TS\w* extends the match to the full
				// alphanumeric run around it, and the optional leading +/- keeps intensity
				// prefixes attached.
				if (Regex.IsMatch(token, "TS")) { tsEntries.Add(new FragEntry(tokenIdx, HighlightedBlockHtml(token, @"[+\-]?\b\w*TS\w*\b"))); meta.TSRed = true; }
				if (Regex.IsMatch(token, "SN|FZRA|FZDZ")) { snowEntries.Add(new FragEntry(tokenIdx, HighlightedBlockHtml(token, @"[+\-]?\b\w*(?:SN|FZRA|FZDZ)\w*\b"))); meta.SnowRed = true; }

				metas.Add(meta);
				tokenIdx++;
			}

			// Second pass: now that every token's own bounds and red flags are known, compute the
			// real effective window for each red flag — see ComputeWindow's header comment.
			List<TafWindow> visCeilWindows = new List<TafWindow>(), windWindows = new List<TafWindow>(),
				tsWindows = new List<TafWindow>(), snowWindows = new List<TafWindow>();
			for (int i = 0; i < metas.Count; i++)
			{
				if (metas[i].VisCeilRed) visCeilWindows.Add(ComputeWindow(metas, i, validFrom, validTo, delegate(TafTokenMeta m) { return m.VisCeilRed; }));
				if (metas[i].WindRed) windWindows.Add(ComputeWindow(metas, i, validFrom, validTo, delegate(TafTokenMeta m) { return m.WindRed; }));
				if (metas[i].TSRed) tsWindows.Add(ComputeWindow(metas, i, validFrom, validTo, delegate(TafTokenMeta m) { return m.TSRed; }));
				if (metas[i].SnowRed) snowWindows.Add(ComputeWindow(metas, i, validFrom, validTo, delegate(TafTokenMeta m) { return m.SnowRed; }));
			}

			// Third pass — drop from the DISPLAYED text (not from the windows/red-flags above,
			// which Check vs Schedule's own 24h+block-hours check already keeps honest) whatever
			// is no longer worth showing as of right now:
			//  - a closed TEMPO/PROB block whose own window has already fully ended;
			//  - a persisting (BECMG/FM/base) block once superseded by a LATER persisting block
			//    that has itself already completed (its own transition end is in the past) AND
			//    clears every one of Ceiling/Vis and Wind (a real, accomplished return to normal
			//    — a still-red or still-amber next block means the earlier one stays relevant as
			//    context for how the trend got here).
			DateTime nowUtc = DateTime.UtcNow;
			bool[] suppressPast = new bool[metas.Count];
			bool[] supersededClear = new bool[metas.Count];
			for (int i = 0; i < metas.Count; i++)
			{
				if (!metas[i].IsPersisting) { suppressPast[i] = metas[i].OwnEnd < nowUtc; continue; }

				int next = -1;
				for (int j = i + 1; j < metas.Count; j++)
					if (metas[j].IsPersisting) { next = j; break; }
				if (next >= 0 && metas[next].OwnEnd < nowUtc && !metas[next].VisCeilRed && !metas[next].WindRed)
					supersededClear[i] = true;
			}

			string visCeilHtml = JoinSurviving(visCeilEntries, suppressPast, supersededClear);
			string windHtml = JoinSurviving(windEntries, suppressPast, supersededClear);
			// TS/Snow only drop past TEMPO/PROB blocks — the "superseded by a cleared BECMG" rule
			// is scoped to Ceiling/Vis and Wind only, per the reasoning above.
			string tsHtml = JoinSurviving(tsEntries, suppressPast, null);
			string snowHtml = JoinSurviving(snowEntries, suppressPast, null);

			TafFragments frag;
			frag.VisCeil = visCeilHtml;
			frag.Wind = windHtml;
			frag.TS = tsHtml;
			frag.Snow = snowHtml;
			frag.VisCeilWindows = visCeilWindows;
			frag.WindWindows = windWindows;
			frag.TSWindows = tsWindows;
			frag.SnowWindows = snowWindows;
			return frag;
		}

		// One category's colored contribution from a single token, tagged with that token's
		// index so a later pass can drop it without needing to re-parse the HTML.
		private struct FragEntry
		{
			public int TokenIndex;
			public string Html;
			public FragEntry(int tokenIndex, string html) { TokenIndex = tokenIndex; Html = html; }
		}

		// Concatenates every entry not suppressed by either filter (supersededClear may be null
		// for categories that rule doesn't apply to).
		private static string JoinSurviving(List<FragEntry> entries, bool[] suppressPast, bool[] supersededClear)
		{
			StringBuilder sb = new StringBuilder();
			foreach (FragEntry e in entries)
			{
				if (suppressPast[e.TokenIndex]) continue;
				if (supersededClear != null && supersededClear[e.TokenIndex]) continue;
				sb.Append(e.Html);
			}
			return sb.ToString();
		}

		// Per-token bookkeeping for the two-pass window computation below.
		private struct TafTokenMeta
		{
			public bool IsPersisting;
			public DateTime OwnStart, OwnEnd;
			public bool VisCeilRed, WindRed, TSRed, SnowRed;
		}

		// A token's OWN transition/closed bounds, straight from its own ddHH/ddHH (BECMG/TEMPO/
		// PROB30/PROB40 all share this literal pattern) or FMddHHmm marker — independent of
		// whether the token is persisting or closed; that distinction is applied afterward by
		// ComputeWindow. Falls back to the whole validity window for the base/initial group (no
		// marker at all) or any unrecognised format — same "unparseable -> fall back to the wider
		// window, never silently hide a real breach" principle as the NOTAM Conflict tab's own
		// schedule parser (ParseNotamActiveWindows, MainForm.Conflict.cs).
		private static void ParseOwnBounds(string token, DateTime validFrom, DateTime validTo, out DateTime start, out DateTime end)
		{
			start = validFrom;
			end = validTo;

			Match range = Regex.Match(token, @"\b(\d{2})(\d{2})/(\d{2})(\d{2})\b");
			if (range.Success)
			{
				DateTime? s = ParseDdHH(range.Groups[1].Value, range.Groups[2].Value, validFrom);
				DateTime? e = ParseDdHH(range.Groups[3].Value, range.Groups[4].Value, validFrom);
				if (s.HasValue) start = s.Value;
				if (e.HasValue) end = e.Value;
				if (end < start) end = end.AddDays(1); // short range wrapping past midnight
				return;
			}
			Match fm = Regex.Match(token, @"FM(\d{2})(\d{2})(\d{2})");
			if (fm.Success)
			{
				DateTime? s = ParseDdHH(fm.Groups[1].Value, fm.Groups[2].Value, validFrom);
				if (s.HasValue) start = s.Value;
				// end left at validTo here — only used if ComputeWindow ever treated this token as
				// non-persisting, which never happens for a real FM marker.
			}
		}

		// A block's effective real-world window:
		// - Non-persisting (TEMPO/PROB30/PROB40): a closed window, exactly its own parsed bounds —
		//   "temporarily from X to Y, then back to whatever the surrounding conditions already
		//   were", per definition. No extension.
		// - Persisting (BECMG/FM/base group): starts at its own start (the earliest the change
		//   could occur — conservative), and runs until the NEXT persisting group (skipping over
		//   any TEMPO/PROB in between, which don't represent a real change of the base state):
		//     - if that next group is ALSO flagged red for this same category (conditions staying
		//       bad or worsening further), cut the current block off at the next one's OWN START —
		//       the two windows sit back-to-back with no gap or overlap, since the next block's
		//       own window already starts there too;
		//     - if the next group is NOT flagged red (i.e. an improvement), conservatively assume
		//       the bad conditions could persist all the way to the next group's OWN END (the
		//       latest point the transition could still be completing);
		//     - if there's no next persisting group at all, extend 24h forward from this block's
		//       own start — deliberately NOT capped at the TAF's own validity end: the TAF simply
		//       has nothing more to say past that point, which is not the same as "conditions are
		//       known to improve" — a dispatcher reading "still bad, nothing else forecast" treats
		//       that as persisting, not as clearing.
		private static TafWindow ComputeWindow(List<TafTokenMeta> metas, int idx, DateTime validFrom, DateTime validTo, Func<TafTokenMeta, bool> redSelector)
		{
			TafTokenMeta meta = metas[idx];
			if (!meta.IsPersisting)
			{
				TafWindow closed;
				closed.Start = meta.OwnStart;
				closed.End = meta.OwnEnd;
				return closed;
			}

			int next = -1;
			for (int j = idx + 1; j < metas.Count; j++)
				if (metas[j].IsPersisting) { next = j; break; }

			TafWindow w;
			w.Start = meta.OwnStart;
			if (next < 0)
			{
				w.End = meta.OwnStart.AddHours(24);
			}
			else
			{
				w.End = redSelector(metas[next]) ? metas[next].OwnStart : metas[next].OwnEnd;
			}
			if (w.End < w.Start) w.End = w.Start;
			return w;
		}

		// "dd" (day of month) + "HH" (hour, 00-23, or 24 meaning next-day 00) -> a UTC DateTime,
		// anchored to referenceUtc's month/year and rolled a month either way if the parsed day
		// ends up more than 15 days from the reference (handles the TAF period crossing a month
		// boundary, e.g. issued the 31st for a period starting "0100").
		private static DateTime? ParseDdHH(string dayStr, string hourStr, DateTime referenceUtc)
		{
			int day, hour;
			if (!int.TryParse(dayStr, out day) || !int.TryParse(hourStr, out hour)) return null;
			if (day < 1 || day > 31) return null;
			if (hour == 24) hour = 0;
			if (hour < 0 || hour > 23) return null;

			DateTime candidate;
			try { candidate = new DateTime(referenceUtc.Year, referenceUtc.Month, 1, hour, 0, 0, DateTimeKind.Utc).AddDays(day - 1); }
			catch { return null; }
			if ((referenceUtc - candidate).TotalDays > 15) candidate = candidate.AddMonths(1);
			else if ((candidate - referenceUtc).TotalDays > 15) candidate = candidate.AddMonths(-1);
			return candidate;
		}

		// The whole trend-group block/clause wrapped in a color span — red (CAT I), amber
		// (advisory tier), or blue (trend recovery); every Vis/Ceiling/Wind caller passes one.
		private static string BlockHtml(string block, string color)
		{
			return "<span style=\"color:" + color + (color == "red" ? ";font-weight:bold\">" : "\">") + block + "</span>";
		}

		// TS/Snow: keep the whole block/clause visible (same context-preserving reasoning as
		// BlockHtml), but only the matched keyword itself (TS/TSRA/SN/FZRA/FZDZ/...) is colored
		// red — the rest of the block stays plain text.
		private static string HighlightedBlockHtml(string block, string keywordPattern)
		{
			return Regex.Replace(block, keywordPattern, delegate(Match m)
			{
				return "<span style=\"color:red;font-weight:bold\">" + m.Value + "</span>";
			});
		}

		// Bolds the speed/gust portion of a matched wind group within its containing token/block
		// — everything after the 3-char direction (e.g. "020" in "02035KT"/"02025G35KT"), so
		// "35KT"/"25G35KT" ends up bold while the direction stays plain weight. windMatch is the
		// match against the ORIGINAL (unmodified) token, so its Index/Length are still valid here.
		private static string BoldWindSpeed(string token, Match windMatch)
		{
			string w = windMatch.Value;
			if (w.Length <= 3) return token;
			string dir = w.Substring(0, 3);
			string speedPart = w.Substring(3);
			string bolded = dir + "<b>" + speedPart + "</b>";
			return token.Substring(0, windMatch.Index) + bolded + token.Substring(windMatch.Index + windMatch.Length);
		}

		// ── Attach image (inserted directly into the editable report) ──

		private void AttachImageToTafReport()
		{
			if (!EnsureWriterOrWarn()) return;
			using (OpenFileDialog dlg = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.bmp", Title = "Attach image to TAF report" })
			{
				if (dlg.ShowDialog() != DialogResult.OK) return;
				try
				{
					// Copied into %TEMP% under a unique name so a since-moved/deleted/overwritten
					// source file can't break re-rendering the saved report later.
					string ext = Path.GetExtension(dlg.FileName);
					string tempPath = Path.Combine(Path.GetTempPath(), "taf_img_" + Guid.NewGuid().ToString("N") + ext);
					File.Copy(dlg.FileName, tempPath, true);
					string src = new Uri(tempPath).AbsoluteUri;
					string imgTag = "<br /><img src=\"" + src + "\" style=\"max-width:600px\" />";

					if (Web_TafAnalysis.Document != null && Web_TafAnalysis.Document.Body != null)
					{
						Web_TafAnalysis.Document.Body.InnerHtml = Web_TafAnalysis.Document.Body.InnerHtml + imgTag;
						SaveEditedTafReport();
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Could not attach image:\n" + ex.Message, "Attach Image", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		// ── Recipients dialog (clone of MainForm.Conflict.cs's ShowRecipientsDialog, targeting
		// TafRecipients instead of EmailRecipients) ──

		private void ShowTafRecipientsDialog()
		{
			using (Form dlg = new Form
			{
				Text = "Dispatch Watch", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false, Width = 440, Height = 480, BackColor = Color.FromArgb(38, 50, 56)
			})
			{
				Label title = new Label { Text = "TAF recipients", ForeColor = Color.White,
					Font = new Font("Segoe UI", 11f, FontStyle.Bold), Top = 14, Left = 20, AutoSize = true };
				dlg.Controls.Add(title);

				ListBox list = new ListBox { Top = 48, Left = 20, Width = 380, Height = 300, BackColor = Color.White, ForeColor = Color.Black };
				foreach (string a in LoadTafRecipients()) list.Items.Add(a);
				dlg.Controls.Add(list);

				TextBox addBox = new TextBox { Top = 358, Left = 20, Width = 270, BackColor = Color.White, ForeColor = Color.Black };
				dlg.Controls.Add(addBox);

				Button add = new Button { Text = "Add", Top = 356, Left = 300, Width = 100, Height = 26,
					BackColor = Color.FromArgb(46, 125, 82), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				add.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				dlg.Controls.Add(add);

				Action doAdd = delegate
				{
					string a = addBox.Text.Trim();
					if (a == "" || !a.Contains("@")) return;
					foreach (object it in list.Items)
						if (string.Equals(it.ToString(), a, StringComparison.OrdinalIgnoreCase)) { addBox.Clear(); return; }
					if (!EnsureWriterOrWarn()) return;

					OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
					conn.Open();
					OleDbCommand ins = new OleDbCommand("INSERT INTO TafRecipients ([Email]) VALUES (?)", conn);
					ins.Parameters.AddWithValue("?", a);
					ins.ExecuteNonQuery();
					conn.Close();

					addBox.Clear();
					list.Items.Clear();
					foreach (string r in LoadTafRecipients()) list.Items.Add(r);
				};
				add.Click += delegate { doAdd(); };
				addBox.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { doAdd(); e.SuppressKeyPress = true; } };

				Button remove = new Button { Text = "Remove selected", Top = 392, Left = 20, Width = 380, Height = 28,
					BackColor = Color.FromArgb(200, 40, 40), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				remove.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				remove.Click += delegate
				{
					if (list.SelectedItem == null) return;
					string a = list.SelectedItem.ToString();
					if (!EnsureWriterOrWarn()) return;

					OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
					conn.Open();
					OleDbCommand del = new OleDbCommand("DELETE FROM TafRecipients WHERE Email=?", conn);
					del.Parameters.AddWithValue("?", a);
					del.ExecuteNonQuery();
					conn.Close();

					list.Items.Clear();
					foreach (string r in LoadTafRecipients()) list.Items.Add(r);
				};
				dlg.Controls.Add(remove);

				Button close = new Button { Text = "Close", Top = 430, Left = 20, Width = 380, Height = 28,
					BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				close.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				close.Click += delegate { dlg.Close(); };
				dlg.Controls.Add(close);

				dlg.ShowDialog();
			}
		}

		// ── Send (clone of MainForm.Email.cs's Btn_sendReportsClick's Outlook-COM mechanics,
		// simplified: no PDF attachments, body = the current TafReport HTML) ──

		void Btn_sendTafReportClick(object sender, EventArgs e)
		{
			if (!EnsureWriterOrWarn()) return;
			SaveEditedTafReport();

			string timeIssued;
			string bodyHtml = LoadTafReportHtml(out timeIssued);
			if (string.IsNullOrEmpty(bodyHtml) || bodyHtml.IndexOf("No TAF analysis yet", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				MessageBox.Show("No TAF analysis to send yet. Click \"TAF Analysis\" first.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			List<string> rcp = LoadTafRecipients();
			if (rcp.Count == 0)
			{
				MessageBox.Show("No TAF recipients defined. Add at least one address.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// A stored TafReport can be non-empty and still stale — it might be left over from a
			// previous session (or a previous dispatcher, on a shared V: copy) rather than a
			// fresh analysis of today's TAFs. Warn explicitly if "TAF Analysis" hasn't actually
			// been clicked yet during this app session, separate from the emptiness check above.
			if (!_tafAnalysisRunThisSession)
			{
				if (MessageBox.Show("TAF Analysis has not been run during this session. The report being sent may be outdated.\n\nSend anyway?",
					"Send TAF Report", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
			}

			if (MessageBox.Show("Send the current TAF report to " + rcp.Count + " recipient(s)?",
				"Send TAF Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

			string titleDate = DateTime.Now.ToString("ddMMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToUpper();
			string subject = "TAF Analysis - " + titleDate;

			// Email mode: raster diagrams via cid: attachments (Outlook's Word engine renders
			// neither VML nor data-URI images in an HTML body — only cid:-attached ones), same
			// dual-mode precedent as BuildConflictReportHtml/Btn_sendReportsClick.
			Dictionary<string, string> inlineImages;
			string checkHtml = BuildTafCheckVsScheduleHtml(true, true, out inlineImages);
			string fullBody = "<html><head><style>" + TafCheckVsScheduleCss() + "</style></head>" +
				"<body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:13px\">" +
				checkHtml + bodyHtml + "</body></html>";

			string step = "init";
			try
			{
				Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
				if (outlookType == null)
				{
					MessageBox.Show("Outlook is not installed or registered on this machine.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				step = "CreateInstance";
				object outlook = Activator.CreateInstance(outlookType);
				step = "CreateItem";
				object mail = outlookType.InvokeMember("CreateItem", BindingFlags.InvokeMethod, null, outlook, new object[] { 0 }); // 0 = olMailItem
				Type mt = mail.GetType();
				step = "Subject";
				mt.InvokeMember("Subject", BindingFlags.SetProperty, null, mail, new object[] { subject });

				step = "Recipients";
				object recips = mt.InvokeMember("Recipients", BindingFlags.GetProperty, null, mail, null);
				Type rct = recips.GetType();
				foreach (string addr in rcp)
				{
					object r = rct.InvokeMember("Add", BindingFlags.InvokeMethod, null, recips, new object[] { addr });
					try { r.GetType().InvokeMember("Type", BindingFlags.SetProperty, null, r, new object[] { 1 }); } catch { } // 1 = olTo
				}
				bool allResolved = (bool)rct.InvokeMember("ResolveAll", BindingFlags.InvokeMethod, null, recips, null);
				if (!allResolved)
				{
					List<string> bad = new List<string>();
					int rcount = (int)rct.InvokeMember("Count", BindingFlags.GetProperty, null, recips, null);
					for (int i = 1; i <= rcount; i++)
					{
						object r = rct.InvokeMember("Item", BindingFlags.GetProperty, null, recips, new object[] { i });
						Type rrt = r.GetType();
						bool res = (bool)rrt.InvokeMember("Resolved", BindingFlags.GetProperty, null, r, null);
						if (!res) bad.Add((string)rrt.InvokeMember("Name", BindingFlags.GetProperty, null, r, null));
					}
					MessageBox.Show("Outlook could not recognise these address(es):\n\n" + string.Join("\n", bad.ToArray()) +
						"\n\nFix or remove them in the TAF recipients list, then try again.",
						"Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				step = "Body";
				int bodyCloseIdx = fullBody.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
				string tail = ReadDefaultSignatureHtml();
				fullBody = bodyCloseIdx >= 0 ? fullBody.Insert(bodyCloseIdx, tail) : fullBody + tail;
				mt.InvokeMember("HTMLBody", BindingFlags.SetProperty, null, mail, new object[] { fullBody });

				// Runway-diagram PNGs for the Check vs Schedule cards: attached as hidden files
				// whose PR_ATTACH_CONTENT_ID matches the "cid:..." reference embedded in the body
				// above — same mechanics as Btn_sendReportsClick (MainForm.Email.cs).
				step = "InlineImages";
				object atts = mt.InvokeMember("Attachments", BindingFlags.GetProperty, null, mail, null);
				Type at = atts.GetType();
				foreach (KeyValuePair<string, string> kv in inlineImages)
				{
					object img = at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { kv.Value });
					Type imgType = img.GetType();
					object pa = imgType.InvokeMember("PropertyAccessor", BindingFlags.GetProperty, null, img, null);
					Type pat = pa.GetType();
					pat.InvokeMember("SetProperty", BindingFlags.InvokeMethod, null, pa,
						new object[] { "http://schemas.microsoft.com/mapi/proptag/0x3712001E", kv.Key });
					pat.InvokeMember("SetProperty", BindingFlags.InvokeMethod, null, pa,
						new object[] { "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", true });
				}

				step = "Send";
				mt.InvokeMember("Send", BindingFlags.InvokeMethod, null, mail, null);

				MessageBox.Show("TAF report sent to " + rcp.Count + " recipient(s).", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				string msg = (ex.InnerException != null) ? ex.InnerException.Message : ex.Message;
				MessageBox.Show("Failed to send via Outlook (step: " + step + "):\n" + msg, "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				foreach (string tempPath in inlineImages.Values)
					try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
			}
		}
	}
}
