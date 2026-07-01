using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Filter-tab pending selections (persisted only on SUBMIT)
		private Dictionary<int, CheckBox[]> _pendImpactChks = new Dictionary<int, CheckBox[]>(); // [A,C,N,D,F,M,R]
		private Dictionary<int, CheckBox>   _pendSupChk     = new Dictionary<int, CheckBox>();
		private Dictionary<int, TextBox>    _pendRemark     = new Dictionary<int, TextBox>();
		private Dictionary<int, TextBox>    _pendSupRemark  = new Dictionary<int, TextBox>();
		private Dictionary<int, string>     _pendRemarkDefault = new Dictionary<int, string>();
		private System.Collections.Generic.HashSet<int> _autoKeepSkip = new System.Collections.Generic.HashSet<int>();
		// NOTAM IDs the engine itself promoted to Status='K' (this session). Only these are
		// ever candidates for auto-demotion — a manual Keep click (Keep_Notam) never adds to
		// this set, so it can never be silently reverted by a later rule re-evaluation.
		private System.Collections.Generic.HashSet<int> _autoKeptIds = new System.Collections.Generic.HashSet<int>();
		// false = auto "Filter New NOTAMS" mode; true = manual ICAO station view (both on tabPage1)
		private bool _stationMode = false;
		private static readonly string[] _impactOrder = { "A", "C", "N", "D", "F", "M", "R" };
		// Light tint of a colour (~82% toward white) for use as a readable background
		private static Color Tint(Color k)
		{
			return Color.FromArgb(
				k.R + (255 - k.R) * 82 / 100,
				k.G + (255 - k.G) * 82 / 100,
				k.B + (255 - k.B) * 82 / 100);
		}

		// A "chip" is a Panel holding a left accent strip + a borderless CheckBox.
		// The checkbox's Tag stores its accent Panel so StyleChk can recolour it.
		private CheckBox CreateChip(Control parent, string label, int left, int top, int width, bool chked, string code)
		{
			Panel chip = new Panel { Tag = "dispose", Left = left, Top = top, Width = width, Height = 22, BackColor = Color.White };
			Panel accent = new Panel { Left = 0, Top = 0, Width = 3, Height = 22, BackColor = Color.Silver };
			CheckBox cb = new CheckBox
			{
				Left = 7, Top = 2, Width = width - 9, Height = 18, Text = label,
				Checked = chked, BackColor = Color.Transparent, Font = new Font("Segoe UI", 8f)
			};
			cb.Tag = accent;
			chip.Controls.Add(accent);
			chip.Controls.Add(cb);
			parent.Controls.Add(chip);
			StyleChk(cb, code);
			return cb;
		}

		// Colour a chip by its impact colour when selected; revert to neutral when not.
		private static void StyleChk(CheckBox c, string code)
		{
			Panel accent = c.Tag as Panel;
			if (c.Checked)
			{
				Color k = ImpactColor(code);
				if (c.Parent != null) c.Parent.BackColor = Tint(k);
				c.ForeColor = k;
				if (accent != null) accent.BackColor = k;
			}
			else
			{
				if (c.Parent != null) c.Parent.BackColor = Color.White;
				c.ForeColor = SystemColors.ControlText;
				if (accent != null) accent.BackColor = Color.Silver;
			}
		}

		// Default keyword seed (used to create/seed the Keywords table on first run).
		private static readonly string[] _defaultKeywords = {
			"CLSD", "U/S", "UNSERVICEABLE", "OUT OF SERVICE",
			"ILS", "GP", "LOC", "RWY", "TWY", "APCH", "DEP",
			"FUEL", "AVBL", "NOT AVBL", "NIL", "LTD",
			"CAT I", "CAT II", "CAT III", "PERM", "H24", "DAILY", "SUP"
		};
		// Live list, loaded from the Keywords table at startup (see LoadKeywords()).
		private static string[] _notamKeywords = _defaultKeywords;

		private static Color ImpactColor(string impact)
		{
			switch (impact)
			{
				case "A":  return Color.FromArgb(200, 40,  40);   // APT CLSD - red
				case "C":  return Color.FromArgb( 80, 80,  80);   // CAT I - gray
				case "N":  return Color.FromArgb(  0, 60, 180);   // No ILS - blue
				case "D":  return Color.FromArgb(200, 90,   0);   // Not ALTN - orangered
				case "F":  return Color.FromArgb(150, 80,   0);   // Fuel - amber
				case "M":  return Color.FromArgb(100,  0, 150);   // MISC - purple
				case "R":  return Color.FromArgb(170, 100,  0);   // RWY - darkorange
				case "AS": return Color.FromArgb(  0, 120, 60);   // AIP SUP - green
				default:   return Color.FromArgb( 80, 80,  80);
			}
		}

		private static string ImpactLabel(string impact)
		{
			switch (impact)
			{
				case "A":  return "AP CLSD";
				case "R":  return "RWY CLSD";
				case "C":  return "CAT I";
				case "N":  return "No ILS";
				case "D":  return "Not ALTN";
				case "F":  return "FUEL";
				case "AS": return "AIP SUP";
				case "M":  return "MISC";
				default:   return "";
			}
		}

		private static void AppendRtb(RichTextBox rtb, string text, Color color, bool bold, float size = 10f)
		{
			rtb.SelectionStart  = rtb.TextLength;
			rtb.SelectionLength = 0;
			rtb.SelectionColor  = color;
			rtb.SelectionFont   = new Font("Courier New", size, bold ? FontStyle.Bold : FontStyle.Regular);
			rtb.AppendText(text);
		}

		private static void HighlightKeywords(RichTextBox rtb)
		{
			HighlightKeywords(rtb, 0);
		}

		private static void HighlightKeywords(RichTextBox rtb, int startChar)
		{
			Font boldFont = new Font("Courier New", 10, FontStyle.Bold);
			Color kwColor = Color.FromArgb(180, 0, 0);
			string txt = rtb.Text;
			foreach (string kw in _notamKeywords)
			{
				int idx = startChar;
				while (true)
				{
					idx = txt.IndexOf(kw, idx, StringComparison.OrdinalIgnoreCase);
					if (idx < 0) break;
					if (IsWholeWord(txt, idx, kw.Length))
					{
						rtb.SelectionStart  = idx;
						rtb.SelectionLength = kw.Length;
						rtb.SelectionFont   = boldFont;
						rtb.SelectionColor  = kwColor;
					}
					idx += kw.Length;
				}
			}
			rtb.SelectionStart  = 0;
			rtb.SelectionLength = 0;
		}

		private static bool IsWordChar(char c)
		{
			return char.IsLetterOrDigit(c);
		}

		// Default remark for an impact: the NOTAM's first line, but if that line is just
		// "No" (scope field) use the validity period "start - end" instead.
		private string NotamRemarkDefault(string text, string fromDate, string tillDate)
		{
			string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
			string first = "";
			foreach (string raw in lines)
			{
				string line = raw.Trim();
				if (line != "") { first = line; break; }
			}
			if (first == "" || first.Equals("No", StringComparison.OrdinalIgnoreCase))
				return FormatDate(fromDate) + " - " + FormatDate(tillDate);
			return first;
		}

		// True if [start, start+len) is bounded by non-alphanumeric chars (whole word).
		// e.g. "LOC" inside "LOCAL" is rejected (followed by 'A').
		private static bool IsWholeWord(string text, int start, int len)
		{
			bool okBefore = start == 0 || !IsWordChar(text[start - 1]);
			int after = start + len;
			bool okAfter = after >= text.Length || !IsWordChar(text[after]);
			return okBefore && okAfter;
		}

		// Builds a VML airport diagram (WebBrowser control runs IE7 mode -> no inline SVG).
		// Each physical runway is a strip oriented by its QFU (designator x 10 deg),
		// scaled by its length, with both threshold labels.
		private static string BuildRwySvg(System.Collections.Generic.List<string> rwyClean)
		{
			int W = 130, H = 110, cx = 65, cy = 55, maxHalf = 40;

			System.Collections.Generic.List<int>    headings = new System.Collections.Generic.List<int>();
			System.Collections.Generic.List<double> lengths  = new System.Collections.Generic.List<double>();
			System.Collections.Generic.List<string> end1     = new System.Collections.Generic.List<string>();
			System.Collections.Generic.List<string> end2     = new System.Collections.Generic.List<string>();

			for (int i = 0; i < rwyClean.Count; i += 2)
			{
				string d1 = ParseDesignator(rwyClean[i]);
				string d2 = (i + 1 < rwyClean.Count) ? ParseDesignator(rwyClean[i + 1]) : "";
				int hdg = ParseHeading(d1);
				if (hdg < 0) continue;
				double len = ParseLength(rwyClean[i]);
				headings.Add(hdg);
				lengths.Add(len);
				end1.Add(d1);
				end2.Add(d2);
			}

			if (headings.Count == 0) return "";

			double maxLen = 0;
			foreach (double l in lengths) if (l > maxLen) maxLen = l;
			if (maxLen <= 0) maxLen = 1;

			System.Text.StringBuilder shapes = new System.Text.StringBuilder();
			System.Text.StringBuilder labels = new System.Text.StringBuilder();

			double spacing = 11;
			for (int i = 0; i < headings.Count; i++)
			{
				double half = maxHalf * (lengths[i] > 0 ? lengths[i] / maxLen : 1.0);
				if (half < 12) half = 12;
				double rad = headings[i] * Math.PI / 180.0;
				double dx = Math.Sin(rad) * half;
				double dy = -Math.Cos(rad) * half;

				// Offset parallel runways (same QFU) perpendicular to their axis
				int parallelIdx = 0;
				for (int k = 0; k < i; k++) if (headings[k] == headings[i]) parallelIdx++;
				double offMag = parallelIdx * spacing;
				double ocx = cx + Math.Cos(rad) * offMag;
				double ocy = cy + Math.Sin(rad) * offMag;

				double x1 = ocx - dx, y1 = ocy - dy;
				double x2 = ocx + dx, y2 = ocy + dy;

				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#607d8b\" strokeweight=\"7px\">" +
					"<v:stroke endcap=\"round\"/></v:line>");
				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#cfd8dc\" strokeweight=\"1px\">" +
					"<v:stroke dashstyle=\"dash\"/></v:line>");
				labels.Append(RwyLabel(end1[i], x1, y1, ocx, ocy));
				if (end2[i] != "") labels.Append(RwyLabel(end2[i], x2, y2, ocx, ocy));
			}

			return "<div style=\"position:relative;width:" + W + "px;height:" + H + "px\">" +
				shapes.ToString() + labels.ToString() + "</div>";
		}

		private static string RwyLabel(string text, double x, double y, double cx, double cy)
		{
			double ox = (x - cx) * 0.22, oy = (y - cy) * 0.22;
			double lx = x + ox - 12, ly = y + oy - 9;
			return "<div style=\"position:absolute;left:" + F(lx) + "px;top:" + F(ly) +
				"px;width:24px;text-align:center;font-size:13px;font-weight:bold;color:#b0bec5;font-family:monospace\">" + text + "</div>";
		}

		private static string F(double v) { return v.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture); }

		// Geographic runway data (from the Runways table) for an accurate diagram.
		private struct RwyGeo { public string Qfu; public double Hdg; public int DistM; public double Lat; public double Lon; }

		private static System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<RwyGeo>> _geoCache =
			new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<RwyGeo>>(StringComparer.OrdinalIgnoreCase);

		private System.Collections.Generic.List<RwyGeo> LoadRwyGeo(string icao)
		{
			if (_geoCache.ContainsKey(icao)) return _geoCache[icao];
			System.Collections.Generic.List<RwyGeo> list = new System.Collections.Generic.List<RwyGeo>();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
				conn.Open();
				OleDbCommand q = new OleDbCommand("SELECT QFU, Hdg, DistM, ThrLat, ThrLon FROM Runways WHERE ICAO=? ORDER BY Ord", conn);
				q.Parameters.AddWithValue("?", icao);
				OleDbDataReader r = q.ExecuteReader();
				while (r.Read())
				{
					RwyGeo g = new RwyGeo();
					g.Qfu   = !r.IsDBNull(0) ? r.GetString(0) : "";
					g.Hdg   = !r.IsDBNull(1) ? Convert.ToDouble(r.GetValue(1)) : 0;
					g.DistM = !r.IsDBNull(2) ? Convert.ToInt32(r.GetValue(2))  : 0;
					g.Lat   = !r.IsDBNull(3) ? Convert.ToDouble(r.GetValue(3)) : 0;
					g.Lon   = !r.IsDBNull(4) ? Convert.ToDouble(r.GetValue(4)) : 0;
					list.Add(g);
				}
				r.Close(); conn.Close();
			}
			catch { }
			// No structured rows with coordinates -> fall back to the in-memory CSV index
			// so the diagram works for ANY ICAO (e.g. KJFK), not just saved stations.
			if (!HasGeo(list)) list = CsvGeoFor(icao);
			// Cache once the result is usable (has at least one located runway). An empty
			// result — e.g. a lookup made before the background CSV index finished loading —
			// is left uncached so the next render retries and picks up the real geo.
			if (HasGeo(list)) _geoCache[icao] = list;
			return list;
		}

		private static bool HasCoords(RwyGeo r) { return r.Lat != 0 || r.Lon != 0; }

		// Usable as soon as at least one threshold has real coordinates. Runways whose
		// thresholds have no coordinates in the open data (e.g. ZSPD 15/33) are simply
		// skipped by BuildRwySvgGeo rather than forcing the whole airport to the schematic.
		private static bool HasGeo(System.Collections.Generic.List<RwyGeo> rs)
		{
			foreach (RwyGeo r in rs) if (HasCoords(r)) return true;
			return false;
		}

		// Airport diagram drawn from the real threshold lat/lon (true crossing angles and
		// parallel spacing). Thresholds are paired by order (le, he, le2, he2, ...). A pair
		// with one missing threshold is projected from heading + length; a pair with no
		// coordinates at all is skipped.
		private static string BuildRwySvgGeo(System.Collections.Generic.List<RwyGeo> rs)
		{
			int W = 130, H = 110, pad = 16;
			int n = rs.Count;
			// Centroid over located thresholds only — coordless (0,0) points must not skew it.
			double lat0 = 0, lon0 = 0; int cnt = 0;
			for (int i = 0; i < n; i++) if (HasCoords(rs[i])) { lat0 += rs[i].Lat; lon0 += rs[i].Lon; cnt++; }
			if (cnt == 0) return "";
			lat0 /= cnt; lon0 /= cnt;
			double cosLat = Math.Cos(lat0 * Math.PI / 180.0);

			System.Collections.Generic.List<double[]> segs = new System.Collections.Generic.List<double[]>();
			System.Collections.Generic.List<object[]> ends = new System.Collections.Generic.List<object[]>();
			for (int i = 0; i < n; i += 2)
			{
				RwyGeo a = rs[i];
				bool hasB = (i + 1 < n);
				RwyGeo b = hasB ? rs[i + 1] : new RwyGeo();
				bool aOk = HasCoords(a), bOk = hasB && HasCoords(b);
				if (!aOk && !bOk) continue;   // whole runway has no coordinates -> skip it

				double ax, ay, bx, by; string aq = a.Qfu, bq = hasB ? b.Qfu : "";
				if (aOk) { ax = (a.Lon - lon0) * cosLat; ay = -(a.Lat - lat0); }
				else     { ax = 0; ay = 0; }
				if (bOk) { bx = (b.Lon - lon0) * cosLat; by = -(b.Lat - lat0); }
				else     { bx = 0; by = 0; }

				if (aOk && !bOk)          // project the far end from a's heading + length
				{
					double L = a.DistM / 111320.0, rad = a.Hdg * Math.PI / 180.0;
					bx = ax + Math.Sin(rad) * L; by = ay - Math.Cos(rad) * L;
					if (!hasB) bq = "";
				}
				else if (bOk && !aOk)     // project the near end back from b's heading + length
				{
					double L = b.DistM / 111320.0, rad = b.Hdg * Math.PI / 180.0;
					ax = bx + Math.Sin(rad) * L; ay = by - Math.Cos(rad) * L;
				}

				segs.Add(new double[] { ax, ay, bx, by });
				ends.Add(new object[] { aq, ax, ay });
				if (bq != "") ends.Add(new object[] { bq, bx, by });
			}
			if (segs.Count == 0) return "";

			double minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
			foreach (double[] s in segs)
			{
				minX = Math.Min(minX, Math.Min(s[0], s[2])); maxX = Math.Max(maxX, Math.Max(s[0], s[2]));
				minY = Math.Min(minY, Math.Min(s[1], s[3])); maxY = Math.Max(maxY, Math.Max(s[1], s[3]));
			}
			double spanX = Math.Max(maxX - minX, 1e-6), spanY = Math.Max(maxY - minY, 1e-6);
			double scale = Math.Min((W - 2.0 * pad) / spanX, (H - 2.0 * pad) / spanY);
			double offX = (W - spanX * scale) / 2.0, offY = (H - spanY * scale) / 2.0;

			System.Text.StringBuilder shapes = new System.Text.StringBuilder();
			System.Text.StringBuilder labels = new System.Text.StringBuilder();
			foreach (double[] s in segs)
			{
				double x1 = offX + (s[0] - minX) * scale, y1 = offY + (s[1] - minY) * scale;
				double x2 = offX + (s[2] - minX) * scale, y2 = offY + (s[3] - minY) * scale;
				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#607d8b\" strokeweight=\"7px\"><v:stroke endcap=\"round\"/></v:line>");
				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#cfd8dc\" strokeweight=\"1px\"><v:stroke dashstyle=\"dash\"/></v:line>");
			}
			foreach (object[] e in ends)
			{
				double x = offX + ((double)e[1] - minX) * scale, y = offY + ((double)e[2] - minY) * scale;
				labels.Append(RwyLabel((string)e[0], x, y, W / 2.0, H / 2.0));
			}
			return "<div style=\"position:relative;width:" + W + "px;height:" + H + "px\">" +
				shapes.ToString() + labels.ToString() + "</div>";
		}

		private static string ParseDesignator(string line)
		{
			int colon = line.IndexOf(':');
			string d = colon > 0 ? line.Substring(0, colon) : line;
			return d.Trim();
		}

		private static int ParseHeading(string designator)
		{
			string digits = "";
			foreach (char c in designator)
			{
				if (c >= '0' && c <= '9') digits += c;
				else break;
			}
			if (digits.Length == 0) return -1;
			int n;
			if (!Int32.TryParse(digits, out n)) return -1;
			return (n * 10) % 360;
		}

		private static double ParseLength(string line)
		{
			// Find a number immediately followed by 'm' (e.g. 3054m)
			for (int i = 0; i < line.Length; i++)
			{
				if ((line[i] == 'm' || line[i] == 'M') && i > 0 && line[i - 1] >= '0' && line[i - 1] <= '9')
				{
					int j = i - 1;
					while (j >= 0 && line[j] >= '0' && line[j] <= '9') j--;
					string num = line.Substring(j + 1, i - j - 1);
					double val;
					if (Double.TryParse(num, System.Globalization.NumberStyles.Integer,
						System.Globalization.CultureInfo.InvariantCulture, out val))
						return val;
				}
			}
			return 0;
		}

		// One-time: add the independent SUP column to filteredNotams_table (idempotent)
		public void EnsureSchema()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("ALTER TABLE filteredNotams_table ADD COLUMN Sup TEXT(3)", conn).ExecuteNonQuery(); }
				catch { /* column already exists */ }
				try { new OleDbCommand("ALTER TABLE filteredNotams_table ADD COLUMN SupRef TEXT(50)", conn).ExecuteNonQuery(); }
				catch { /* column already exists */ }

				// Migrate legacy SUP (Impact='AS', reference stored in Remark) to the new model
				// (Sup='Yes', reference in SupRef). Idempotent: leaves no Impact='AS' rows.
				try { new OleDbCommand("UPDATE filteredNotams_table SET SupRef=Remark WHERE Impact='AS' AND (SupRef IS NULL OR SupRef='') AND Remark<>''", conn).ExecuteNonQuery(); } catch { }
				try { new OleDbCommand("UPDATE filteredNotams_table SET Sup='Yes', Impact='', Remark='' WHERE Impact='AS'", conn).ExecuteNonQuery(); } catch { }
				conn.Close();
			}
			catch { /* DB not ready; ignore */ }
		}

		// ── Auto-classification engine ───────────────────────────────────────
		// Suggests impact checkboxes from NOTAM text + airport runway context.
		// Suggestions are visual only (AUTO state); nothing is written until the
		// dispatcher confirms by clicking the checkbox.
		private struct ImpactSuggestion
		{
			public bool APClsd, CatI, NoILS, NotAltn, Fuel, Sup;
			// Auto-Keep-only signals: true when this NOTAM mentions a runway closure / an ILS
			// outage but doesn't otherwise trigger any of the selectable impact codes above
			// (e.g. a single runway/ILS down with a parallel/other one still available).
			// SuggestedSingleCode ignores these — they only make the auto-Keep pre-pass
			// surface the NOTAM for manual review, since it's always worth a look even
			// without an automatic impact.
			public bool RwyClosure;
			public bool IlsOutage;
		}

		private struct RwyInfo
		{
			public string Desig;   // e.g. "27", "09L"
			public int    CatMax;  // 0 = none, 1/2/3 = ILS CAT level
		}

		private static System.Collections.Generic.List<RwyInfo> ParseRunways(string rwyField)
		{
			System.Collections.Generic.List<RwyInfo> list = new System.Collections.Generic.List<RwyInfo>();
			string[] lines = rwyField.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
			foreach (string raw in lines)
			{
				string line = raw.Trim();
				if (line == "") continue;
				RwyInfo r = new RwyInfo();
				int colon = line.IndexOf(':');
				r.Desig = (colon > 0 ? line.Substring(0, colon) : line).Trim();
				r.CatMax = 0;
				string up = line.ToUpper();
				if (up.Contains("CAT 3") || up.Contains("CAT III")) r.CatMax = 3;
				else if (up.Contains("CAT 2") || up.Contains("CAT II")) r.CatMax = 2;
				else if (up.Contains("CAT 1") || up.Contains("CAT I"))  r.CatMax = 1;
				list.Add(r);
			}
			return list;
		}

		// Extracts the AIP SUP reference, e.g. "SUP 056/2026", from NOTAM text
		private static string ExtractSupRef(string text)
		{
			System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(
				text, @"SUP\s*0*\d+\s*/\s*\d{2,4}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
			if (!m.Success) return "";
			// Normalise to "SUP nnn/yyyy"
			System.Text.RegularExpressions.Match n = System.Text.RegularExpressions.Regex.Match(
				m.Value, @"(\d+)\s*/\s*(\d{2,4})");
			if (!n.Success) return "SUP";
			return "SUP " + n.Groups[1].Value + "/" + n.Groups[2].Value;
		}

		private static bool RegexAny(string text, params string[] patterns)
		{
			foreach (string p in patterns)
				if (System.Text.RegularExpressions.Regex.IsMatch(text, p,
					System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return true;
			return false;
		}

		// Thresholds closed by THIS NOTAM (e.g. "RWY 09/27 CLSD. RWY 08/26 CLSD" -> 09,27,08,26).
		// Only meaningful when the NOTAM actually contains a closure.
		private static System.Collections.Generic.List<string> ClosedThresholds(string U)
		{
			System.Collections.Generic.List<string> set = new System.Collections.Generic.List<string>();
			if (!RegexAny(U, @"\bCLSD", @"\bCLOSED")) return set;
			foreach (System.Text.RegularExpressions.Match m in
				System.Text.RegularExpressions.Regex.Matches(U, @"(\d{1,2}[LCR]?)\s*/\s*(\d{1,2}[LCR]?)"))
			{
				if (!set.Contains(m.Groups[1].Value)) set.Add(m.Groups[1].Value);
				if (!set.Contains(m.Groups[2].Value)) set.Add(m.Groups[2].Value);
			}
			foreach (System.Text.RegularExpressions.Match m in
				System.Text.RegularExpressions.Regex.Matches(U, @"R(?:WY|UNWAY)\s*(\d{1,2}[LCR]?)\b(?!\s*/)"))
				if (!set.Contains(m.Groups[1].Value)) set.Add(m.Groups[1].Value);
			return set;
		}

		// Data-derived rules, validated against the dispatcher's historical classifications
		// (see ICAO_storedNotams.mdb analysis). allKeptUpper kept for signature compatibility.
		private static ImpactSuggestion SuggestImpacts(string notamText,
			System.Collections.Generic.List<RwyInfo> runways,
			System.Collections.Generic.List<string> allKeptUpper)
		{
			ImpactSuggestion s = new ImpactSuggestion();
			string U = notamText.ToUpper();

			// SUP — independent
			s.Sup = U.Contains("SUP");

			// Fuel (3/3 on sample)
			s.Fuel = RegexAny(U, @"FUEL.{0,20}(NOT\s+AVBL|U/S|NIL|UNAVAIL)", @"NO\s+FUEL", @"FUEL\s+DISRUPTION");

			// Not as alternate / PPR / delay (7/7 on sample).
			// \bPPR\b excludes two common false-positive shapes seen on NAVAID/TWY NOTAMs:
			// a directly-following phone number ("PPR 617-561-1919" — a ground-ops contact for
			// taxi/parking access) or a short duration ("AVBL PPR 10MIN" — a prior-notice
			// window for using degraded equipment, not a restriction on using the airport as
			// an alternate).
			s.NotAltn = RegexAny(U, @"\bPPR\b(?!\s*\d[\d\-]{5,})(?!\s*\d{1,3}\s*MIN\b)", @"PRIOR\s+PERMISSION", @"CANNOT\s+BE\s+CHOSEN\s+AS",
				@"NOT.{0,12}ALTERNATE", @"NOT\s+AVBL\s+AS\s+ALTN", @"\bDIVERSION", @"SUBJ.{0,10}DLA",
				@"EXPECT\s+DELAY", @"DELAY\s+EXPECTED", @"\bO/R\s+ONLY", @"NOT\s+AVBL\s+FOR\s+LANDING");

			// LVP exception => ILS available except in low-vis => CAT I, not "No ILS"
			bool lvpExc = RegexAny(U, @"(EXC|EXCEPT)\s+LVP");

			// A DME "associated with" the ILS is a supporting component, not the ILS itself
			// (LOC/GP) — losing it doesn't take the ILS approach out of service.
			bool dmeOnly = RegexAny(U, @"\bDME\b.{0,20}ASSOCIATED\s+WITH\s+ILS.{0,15}(U/S|UNSERVICEABLE|NOT\s+AV(BL|AILABLE))");

			// No ILS (2/2): ILS unserviceable, unless it is only an LVP-only restriction, a
			// DME-only outage (see above), OR another runway at the airport still has a
			// working ILS (mirrors the APT CLSD "all runways closed" check below) — losing
			// the ILS on one runway shouldn't flag "No ILS" airport-wide when a
			// parallel/other runway remains ILS-equipped.
			bool ilsOutageText = !lvpExc && !dmeOnly &&
				RegexAny(U, @"\bILS\b.{0,30}(U/S|UNSERVICEABLE|NOT\s+AV(BL|AILABLE)|NOT\s+USABLE|ON\s+TEST)");
			s.NoILS = ilsOutageText;
			s.IlsOutage = ilsOutageText;
			if (ilsOutageText && runways.Count > 0)
			{
				System.Collections.Generic.List<string> affectedRwys = new System.Collections.Generic.List<string>();
				foreach (System.Text.RegularExpressions.Match m in
					System.Text.RegularExpressions.Regex.Matches(U, @"RWY\s*(\d{1,2}[LCR]?)"))
					if (!affectedRwys.Contains(m.Groups[1].Value)) affectedRwys.Add(m.Groups[1].Value);

				foreach (RwyInfo r in runways)
					if (r.CatMax >= 1 && !affectedRwys.Contains(r.Desig.ToUpper())) { s.NoILS = false; break; }
			}

			// CAT I (3/6): CAT II/III lost, or downgrade to CAT I, or ILS U/S except LVP
			s.CatI = RegexAny(U, @"CAT\s*(II|III|2|3)\b.{0,28}(NOT\s+(AUTH|AVBL|AVAILABLE)|U/S|UNSERVICEABLE|DOWNGRAD)",
				@"DOWNGRAD.{0,15}CAT\s*(I|1)\b")
				|| (lvpExc && RegexAny(U, @"ILS.{0,20}U/S"));

			// APT CLSD (22/30, 4 FP): aerodrome closed, OR this NOTAM closes EVERY runway
			bool apClsdText = RegexAny(U, @"\bAD\s+CLSD", @"\bARP\s+CLSD", @"AERODROME\s+CLOSED", @"AIRPORT\s+CLOSED");
			System.Collections.Generic.List<string> closed = ClosedThresholds(U);
			bool allRwyClosed = false;
			if (runways.Count > 0)
			{
				allRwyClosed = true;
				foreach (RwyInfo r in runways)
					if (!closed.Contains(r.Desig.ToUpper())) { allRwyClosed = false; break; }
			}
			s.APClsd = apClsdText || allRwyClosed;

			// A runway closure that doesn't close the whole airport still deserves review
			// (auto-Keep only — see the RwyClosure comment above).
			s.RwyClosure = closed.Count > 0 && !s.APClsd;

			return s;
		}

		// Shared dark "airport card" HTML (ICAO + IATA + RWY blocks + VML diagram).
		// Used by both the Filter tab header and the Stations tab header.
		private string BuildAirportCardHtml(string AP, string RWYs, int newCount, out int headerHeight)
		{
			string iata = GetIATA(AP);
			string[] rwyLines = RWYs.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
			System.Collections.Generic.List<string> rwyClean = new System.Collections.Generic.List<string>();
			foreach (string rl in rwyLines) if (rl.Trim() != "") rwyClean.Add(rl.Trim());
			int pairRows = (rwyClean.Count + 1) / 2;
			headerHeight = Math.Max(72 + pairRows * 26 + 12, 72 + 130);

			string iataLine = (iata != "" && iata != AP) ? "<div class=\"sub\">IATA: " + iata + "</div>" : "";
			string leftCol = "", rightCol = "";
			for (int i = 0; i < rwyClean.Count; i++)
			{
				string cell = "<div class=\"rwyline\">" +
					rwyClean[i].Replace("&", "&amp;").Replace("<", "&lt;") + "</div>";
				if (i % 2 == 0) leftCol += cell; else rightCol += cell;
			}
			System.Collections.Generic.List<RwyGeo> geo = LoadRwyGeo(AP);
			string rwySvg = HasGeo(geo) ? BuildRwySvgGeo(geo) : "";
			if (rwySvg == "") rwySvg = BuildRwySvg(rwyClean);   // schematic fallback (no geo yet / no coords)
			return
				"<html xmlns:v=\"urn:schemas-microsoft-com:vml\"><head><style>" +
				"v\\:*{behavior:url(#default#VML)}" +
				"body{margin:0;padding:10px 14px;background:#263238;font-family:'Courier New',monospace;overflow:hidden;position:relative}" +
				".icao{font-size:22px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".newbadge{font-size:13px;font-weight:normal;letter-spacing:0;color:#ffca28;margin-left:12px}" +
				".sub{font-size:13px;color:#78909c;margin-top:2px;margin-bottom:10px}" +
				".blk{font-size:13px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:6px 14px;margin-top:10px;margin-right:10px;vertical-align:top}" +
				".rwyline{white-space:nowrap;line-height:1.9}" +
				".diagram{position:absolute;top:8px;right:14px}" +
				"</style></head><body>" +
				"<div class=\"diagram\">" + rwySvg + "</div>" +
				"<div class=\"icao\">" + AP +
				(newCount > 0 ? "<span class=\"newbadge\">" + newCount + " new</span>" : "") +
				"</div>" +
				iataLine +
				"<table cellspacing=\"0\" cellpadding=\"0\"><tr>" +
				"<td class=\"blk\">" + leftCol + "</td>" +
				"<td class=\"blk\">" + rightCol + "</td>" +
				"</tr></table>" +
				"</body></html>";
		}

		// Shared read-only "kept NOTAM" card (colored impact strip + borderless RichTextBox).
		// Returns the card height. Used by the left column of both Filter and Stations tabs.
		private int RenderKeptCard(Control parent, int left, int top, int width,
			string fromDate, string tillDate, string text, string notamKey, string Impact, string Remark)
		{
			Color ic = ImpactColor(Impact);
			string ilabel = ImpactLabel(Impact);
			RichTextBox rtb = new RichTextBox
			{
				Tag = "dispose", Left = 6, Top = 0, Width = width - 8, Height = 2000,
				BorderStyle = BorderStyle.None, ReadOnly = true,
				BackColor = SystemColors.Window, ScrollBars = RichTextBoxScrollBars.None
			};
			AppendRtb(rtb, notamKey, ic, true, 11f);
			if (ilabel != "") AppendRtb(rtb, "  [" + ilabel + "]", ic, false, 9f);
			AppendRtb(rtb, "\n", Color.Black, false);
			AppendRtb(rtb, FormatDate(fromDate) + "  →  " + FormatDate(tillDate) + "\n", Color.DimGray, false, 9f);
			int textStart = rtb.TextLength;
			AppendRtb(rtb, text + "\n", Color.Black, false, 10f);
			if (Remark != "") AppendRtb(rtb, "▶ " + Remark + "\n", ic, false, 9f);
			HighlightKeywords(rtb, textStart);

			int totalLines = rtb.GetLineFromCharIndex(rtb.TextLength) + 1;
			int rtbHeight  = totalLines * 17 + 12;
			rtb.Height = rtbHeight;

			Panel container = new Panel { Tag = "dispose", Left = left, Top = top, Width = width, Height = rtbHeight };
			Panel strip     = new Panel { Left = 0, Top = 0, Width = 4, Height = rtbHeight, BackColor = ic };
			container.Controls.Add(strip);
			container.Controls.Add(rtb);
			parent.Controls.Add(container);
			return rtbHeight;
		}

		void Filter_Notams()
		{
			// Must happen before anything below is positioned: adding/positioning controls
			// while the AutoScroll panel is still scrolled bakes the current scroll offset
			// into their effective location, so a later reset can't undo it.
			tabPage1.AutoScrollPosition = new Point(0, 0);
			AlignTopBar();
			_stationMode = false;
			ClearTaggedControls(tabPage1);
			_pendImpactChks.Clear();
			_pendSupChk.Clear();
			_pendRemark.Clear();
			_pendSupRemark.Clear();
			_pendRemarkDefault.Clear();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();

			string AP = "";
			OleDbDataReader APreader = new OleDbCommand(
				"SELECT location FROM filteredNotams_table WHERE (Checked='N') ORDER BY location DESC", conn).ExecuteReader();
			while (APreader.Read())
				if (!APreader.IsDBNull(0)) AP = APreader.GetString(0);
			conn.Close();

			// No unchecked NOTAMs left -> show the "all checked" state and stop
			if (AP == "")
			{
				Web_FilterHeader.Location = new Point(7, 44);
				Web_FilterHeader.Size = new Size(490, 100);
				string ts = DateTime.Now.ToString("dd MMM yyyy HH:mm").ToUpper() + "z";
				Web_FilterHeader.DocumentText =
					"<html><head><style>" +
					"body{margin:0;padding:0;background:#263238;font-family:'Courier New',monospace;overflow:hidden}" +
					".card{padding:22px 28px}" +
					".tick{float:left;width:48px;height:48px;line-height:48px;text-align:center;border-radius:24px;background:#2e7d52;color:#fff;font-size:26px;margin-right:18px}" +
					".title{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:2px;padding-top:4px}" +
					".sub{font-size:11px;color:#78909c;margin-top:3px}" +
					"</style></head><body><div class=\"card\">" +
					"<div class=\"tick\">&#10003;</div>" +
					"<div class=\"title\">ALL NOTAMS CHECKED</div>" +
					"<div class=\"sub\">Last reviewed &middot; " + ts + "</div>" +
					"</div></body></html>";
				Lbl_location.Text        = "";
				Lbl_notamsUnchecked.Text = "Notams Unchecked : 0";
				Btn_submitNotams.Visible = false;
				tabPage1.AutoScrollPosition = new Point(0, 0);
				return;
			}
			Btn_submitNotams.Visible = true;

			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			string RWYs = "";
			while (OCCreader.Read())
				if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			// Count of new (unchecked) NOTAMs for the airport card badge
			conn.Open();
			OleDbCommand cmdCnt = new OleDbCommand(
				"SELECT COUNT(*) FROM filteredNotams_table WHERE (Checked='N') AND (location=?)", conn);
			cmdCnt.Parameters.AddWithValue("?", AP);
			int newCount = Convert.ToInt32(cmdCnt.ExecuteScalar());
			conn.Close();

			// Header WebBrowser: shared dark airport card (below the top bar)
			int headerHeight;
			Web_FilterHeader.Location = new Point(7, 44);
			Web_FilterHeader.DocumentText = BuildAirportCardHtml(AP, RWYs, newCount, out headerHeight);
			Web_FilterHeader.Size = new Size(490, headerHeight);

			// Per-NOTAM RTBs with colored left border strip (Option B)
			System.Collections.Generic.List<RwyInfo> runways = ParseRunways(RWYs);

			// Auto-Keep / Auto-Un-Keep: promote not-yet-kept NOTAMs to Status='K' when the
			// engine detects a potential impact/SUP/closure/ILS outage, and demote NOTAMs
			// that the engine itself auto-Kept earlier (tracked in _autoKeptIds, this session
			// only) if a classification-rule change means they no longer trigger any signal
			// (e.g. a TWY/DVOR/DME NOTAM wrongly auto-Kept by an older, looser rule). Only
			// engine-promoted NOTAMs are ever demoted — a manual Keep click (Keep_Notam) never
			// adds to _autoKeptIds, so it can never be silently reverted here, even before the
			// dispatcher has picked an impact for it. This must run BEFORE the "kept"
			// left-column query below, otherwise a change made on this very render is missed
			// until the next full re-render (e.g. a manual Keep click).
			conn.Open();
			OleDbCommand cmdScan = new OleDbCommand(
				"SELECT ID, [all], Status FROM filteredNotams_table WHERE (Checked='N') AND (location=?)", conn);
			cmdScan.Parameters.AddWithValue("?", AP);
			OleDbDataReader scanR = cmdScan.ExecuteReader();
			System.Collections.Generic.List<int> toKeep = new System.Collections.Generic.List<int>();
			System.Collections.Generic.List<int> toUnKeep = new System.Collections.Generic.List<int>();
			while (scanR.Read())
			{
				int sid = !scanR.IsDBNull(0) ? scanR.GetInt32(0) : 0;
				string status = !scanR.IsDBNull(2) ? scanR.GetString(2) : "";
				bool kept = status == "K";

				if (kept && !_autoKeptIds.Contains(sid)) continue;    // manual Keep — never touch
				if (!kept && _autoKeepSkip.Contains(sid)) continue;   // dispatcher explicitly ignored it

				string stxt = !scanR.IsDBNull(1) ? scanR.GetString(1).Replace("(char)39", "'") : "";
				// keptUpper is unused by SuggestImpacts (rules are per-NOTAM) — safe to pass empty here.
				ImpactSuggestion sg = SuggestImpacts(stxt, runways, new System.Collections.Generic.List<string>());
				bool signal = SuggestedSingleCode(sg) != "" || sg.Sup || sg.RwyClosure || sg.IlsOutage;

				if (!kept && signal) toKeep.Add(sid);
				else if (kept && !signal) toUnKeep.Add(sid);
			}
			scanR.Close();
			foreach (int kid in toKeep)
			{
				OleDbCommand uk = new OleDbCommand("UPDATE filteredNotams_table SET Status='K' WHERE ID=?", conn);
				uk.Parameters.AddWithValue("?", kid);
				uk.ExecuteNonQuery();
				_autoKeptIds.Add(kid);
			}
			foreach (int uid in toUnKeep)
			{
				OleDbCommand un = new OleDbCommand("UPDATE filteredNotams_table SET Status='' WHERE ID=?", conn);
				un.Parameters.AddWithValue("?", uid);
				un.ExecuteNonQuery();
				_autoKeptIds.Remove(uid);
			}
			conn.Close();

			System.Collections.Generic.List<string> keptUpper = new System.Collections.Generic.List<string>();
			int keptTop = Web_FilterHeader.Bottom + 8;
			conn.Open();
			OleDbCommand cmdKept = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)", conn);
			cmdKept.Parameters.AddWithValue("?", AP);
			OleDbDataReader keptReader = cmdKept.ExecuteReader();
			while (keptReader.Read())
			{
				string fromDate = !keptReader.IsDBNull(5)  ? keptReader.GetString(5)  : "";
				string tillDate = !keptReader.IsDBNull(6)  ? keptReader.GetString(6)  : "";
				string text     = !keptReader.IsDBNull(7)  ? keptReader.GetString(7)  : "";
				string notamKey = !keptReader.IsDBNull(10) ? keptReader.GetString(10) : "";
				string Impact   = !keptReader.IsDBNull(13) ? keptReader.GetString(13) : "";
				string Remark   = !keptReader.IsDBNull(14) ? keptReader.GetString(14) : "";
				text = text.Replace("(char)39", "'");
				keptUpper.Add(text.ToUpper());
				keptTop += RenderKeptCard(tabPage1, 7, keptTop, 490, fromDate, tillDate, text, notamKey, Impact, Remark) + 6;
			}
			conn.Close();

			// New unchecked NOTAMs
			FontFamily courier = new FontFamily("Courier New");
			conn.Open();
			OleDbCommand cmdNew = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Checked='N') AND (location=?)", conn);
			cmdNew.Parameters.AddWithValue("?", AP);
			OleDbDataReader dBreader = cmdNew.ExecuteReader();

			int nbNotams = 0;
			int Top = 48;   // below the top bar (Filter New / ICAO search)

			Dictionary<int, Button>      keep_Buttons       = new Dictionary<int, Button>();

			while (dBreader.Read())
			{
				int    notam_ID   = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
				string fromDate   = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate   = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text = !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key  = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status     = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact     = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark     = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				Color keyColor = HasImpact(Impact) ? ImpactColor(Impact) : Color.MidnightBlue;
				Color stripColor = HasImpact(Impact) ? ImpactColor(Impact) : Color.FromArgb(120, 130, 140);

				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				int ctrlH  = 124;
				int cardH  = Math.Max(20 + height, ctrlH) + 10;
				int cardAvail = tabPage1.ClientSize.Width - 505 - SystemInformation.VerticalScrollBarWidth - 12;
				if (cardAvail < 700) cardAvail = 700;

				// Light control-box background (pushed behind the controls later)
				int ctrlLeft = 1064;
				int ctrlW    = cardAvail - (ctrlLeft - 505) - 4;   // ends just inside the card
				Panel ctrlBox = null;
				if (Status == "K")
				{
					ctrlBox = new Panel
					{
						Tag = "dispose", Left = ctrlLeft, Top = Top, Width = ctrlW, Height = ctrlH,
						BackColor = Color.FromArgb(247, 248, 250), BorderStyle = BorderStyle.FixedSingle
					};
					tabPage1.Controls.Add(ctrlBox);
				}

				// White card with impact-coloured left strip (pushed to the very back later)
				Panel card  = new Panel { Tag = "dispose", Left = 505, Top = Top - 4, Width = cardAvail, Height = cardH, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
				Panel cstrip = new Panel { Left = 0, Top = 0, Width = 4, Height = cardH - 2, BackColor = stripColor };
				card.Controls.Add(cstrip);
				tabPage1.Controls.Add(card);

				AddNotamLabel(tabPage1, courier, notam_key, Top+2, 516, 140, keyColor, true, 11f);
				AddNotamLabel(tabPage1, courier, FormatDate(fromDate), Top+2, 655, 190, Color.DimGray, false, 10f);
				AddNotamLabel(tabPage1, courier, FormatDate(tillDate), Top+2, 850, 190, Color.DimGray, false, 10f);

				tabPage1.Controls.Add(MakeNotamWebBrowser(notam_text, Top+22, 510, height, Status == "K"));

				keep_Buttons[notam_ID] = MakeKeepButton(courier, Status, new Point(1070, Top+6));
				int nid = notam_ID;
				if (Status == "")  keep_Buttons[notam_ID].Click += (s, e) => Keep_Notam(nid);
				if (Status == "K") keep_Buttons[notam_ID].Click += (s, e) => Ignore_Notam(nid);
				tabPage1.Controls.Add(keep_Buttons[notam_ID]);

				if (Status == "K")
				{
					bool stored = HasImpact(Impact);
					int supOrd = dBreader.GetOrdinal("Sup");
					int supRefOrd = dBreader.GetOrdinal("SupRef");
					bool supStored = !dBreader.IsDBNull(supOrd) && dBreader.GetString(supOrd) == "Yes";
					string storedSupRef = !dBreader.IsDBNull(supRefOrd) ? dBreader.GetString(supRefOrd) : "";

					// Auto-suggestion only when nothing is stored yet
					ImpactSuggestion sug = SuggestImpacts(notam_text, runways, keptUpper);
					string sugCode = SuggestedSingleCode(sug);

					string remarkDefault = NotamRemarkDefault(notam_text, fromDate, tillDate);
					AddFilterCheckboxes(notam_ID, Impact, stored, sugCode,
						supStored, sug.Sup, notam_text, remarkDefault, Remark, storedSupRef, Top, ctrlLeft, ctrlW);
				}

				// Push the background panels behind all content added above
				if (ctrlBox != null) ctrlBox.SendToBack();
				card.SendToBack();

				Top = Top + cardH + 8;
				nbNotams++;
			}
			conn.Close();

			Lbl_location.Text        = AP;
			Lbl_notamsUnchecked.Text = "Notams Unchecked : " + nbNotams;

			int barTop  = Top + 20;
			int barLeft = 505;
			int barW    = tabPage1.ClientSize.Width - 505 - SystemInformation.VerticalScrollBarWidth - 12;
			if (barW < 700) barW = 700;
			int barH    = 38;

			Panel statusBar = new Panel
			{
				Tag = "dispose", Top = barTop, Left = barLeft,
				Width = barW, Height = barH,
				BackColor = Color.FromArgb(26, 26, 26)
			};
			Label statusLbl = new Label
			{
				Text      = nbNotams + " NOTAM" + (nbNotams > 1 ? "s" : "") + " unchecked  ·  " + AP,
				ForeColor = Color.FromArgb(144, 164, 174),
				Font      = new Font("Courier New", 10),
				AutoSize  = true, Top = 10, Left = 14
			};
			statusBar.Controls.Add(statusLbl);
			tabPage1.Controls.Add(statusBar);
			statusBar.SendToBack();

			Btn_submitNotams.Top  = barTop + 4;
			Btn_submitNotams.Left = barLeft + barW - Btn_submitNotams.Width - 14;
			Btn_submitNotams.BringToFront();

			// Reset the scroll position again now that all content is built. Doing this only
			// at the top of the method (before ClearTaggedControls/rebuild) is unreliable —
			// AutoScrollMinSize is recalculated once the new content is added, which can
			// silently restore whatever the previous scroll offset was. Setting
			// AutoScrollPosition (not just VerticalScroll.Value) here guarantees the top bar,
			// which sits at a fixed Top/Left, is actually visible at the top of the tab.
			tabPage1.AutoScrollPosition = new Point(0, 0);
		}

		void ICAO_Notams()
		{
			// See the comment in Filter_Notams: this must run before anything is positioned.
			tabPage1.AutoScrollPosition = new Point(0, 0);
			AlignTopBar();
			_stationMode = true;
			ClearTaggedControls(tabPage1);
			Btn_submitNotams.Visible = false;   // no SUBMIT in station view

			string AP = TxtBox_ICAO.Text.Trim().ToUpper();

			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			string RWYs = "";
			while (OCCreader.Read())
				if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			// Airport card into the shared header (below the top bar)
			int headerHeight;
			Web_FilterHeader.Location = new Point(7, 44);
			Web_FilterHeader.DocumentText = BuildAirportCardHtml(AP, RWYs, 0, out headerHeight);
			Web_FilterHeader.Size = new Size(490, headerHeight);

			FontFamily courier = new FontFamily("Courier New");
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");

			// LEFT column — Kept NOTAMs, read-only cards below the airport card
			int keptTop = Web_FilterHeader.Bottom + 8;
			conn.Open();
			OleDbCommand cmdKept = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)", conn);
			cmdKept.Parameters.AddWithValue("?", AP);
			OleDbDataReader keptR = cmdKept.ExecuteReader();
			while (keptR.Read())
			{
				string fromDate = !keptR.IsDBNull(5)  ? keptR.GetString(5)  : "";
				string tillDate = !keptR.IsDBNull(6)  ? keptR.GetString(6)  : "";
				string text     = !keptR.IsDBNull(7)  ? keptR.GetString(7)  : "";
				string notamKey = !keptR.IsDBNull(10) ? keptR.GetString(10) : "";
				string Impact   = !keptR.IsDBNull(13) ? keptR.GetString(13) : "";
				string Remark   = !keptR.IsDBNull(14) ? keptR.GetString(14) : "";
				text = text.Replace("(char)39", "'");
				keptTop += RenderKeptCard(tabPage1, 7, keptTop, 490, fromDate, tillDate, text, notamKey, Impact, Remark) + 6;
			}
			conn.Close();

			// RIGHT column — ALL NOTAMs of the station (kept + ignored + new), each a card
			// with a Keep/Ignore button + impact chips (immediate write).
			conn.Open();
			OleDbCommand cmdAll = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (location=?) ORDER BY Status DESC, key", conn);
			cmdAll.Parameters.AddWithValue("?", AP);
			OleDbDataReader dBreader = cmdAll.ExecuteReader();

			int Top = 48;   // below the top bar
			int cardAvail = tabPage1.ClientSize.Width - 505 - SystemInformation.VerticalScrollBarWidth - 12;
			if (cardAvail < 700) cardAvail = 700;

			while (dBreader.Read())
			{
				int    notam_ID   = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
				string fromDate   = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate   = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text = !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key  = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status     = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact     = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark     = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				int    supOrd     = dBreader.GetOrdinal("Sup");
				int    supRefOrd  = dBreader.GetOrdinal("SupRef");
				bool   supStored  = !dBreader.IsDBNull(supOrd) && dBreader.GetString(supOrd) == "Yes";
				string storedSupRef = !dBreader.IsDBNull(supRefOrd) ? dBreader.GetString(supRefOrd) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				bool kept = (Status == "K");
				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				int ctrlH  = 130;
				int cardH  = Math.Max(20 + height, kept ? ctrlH : 90) + 10;
				int ctrlLeft = 1064;
				int ctrlW    = cardAvail - (ctrlLeft - 505) - 4;

				// Control box + impact chips ONLY for Kept NOTAMs (not for ignored/new)
				Panel ctrlBox = null;
				if (kept)
				{
					ctrlBox = new Panel
					{
						Tag = "dispose", Left = ctrlLeft, Top = Top, Width = ctrlW, Height = ctrlH,
						BackColor = Color.FromArgb(247, 248, 250), BorderStyle = BorderStyle.FixedSingle
					};
					tabPage1.Controls.Add(ctrlBox);
				}

				Color stripColor = HasImpact(Impact) ? ImpactColor(Impact) : (kept ? Color.FromArgb(60, 110, 180) : Color.FromArgb(120, 130, 140));
				Panel card  = new Panel { Tag = "dispose", Left = 505, Top = Top - 4, Width = cardAvail, Height = cardH, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
				Panel cstrip = new Panel { Left = 0, Top = 0, Width = 4, Height = cardH - 2, BackColor = stripColor };
				card.Controls.Add(cstrip);
				tabPage1.Controls.Add(card);

				Color keyColor = HasImpact(Impact) ? ImpactColor(Impact) : Color.MidnightBlue;
				AddNotamLabel(tabPage1, courier, notam_key, Top+2, 516, 140, keyColor, true, 11f);
				AddNotamLabel(tabPage1, courier, FormatDate(fromDate), Top+2, 655, 190, Color.DimGray, false, 10f);
				AddNotamLabel(tabPage1, courier, FormatDate(tillDate), Top+2, 850, 190, Color.DimGray, false, 10f);
				tabPage1.Controls.Add(MakeNotamWebBrowser(notam_text, Top+22, 510, height, kept));

				Button keepBtn = MakeKeepButton(courier, Status, new Point(1070, Top+6));
				int nid = notam_ID;
				if (kept) keepBtn.Click += (s, e) => Ignore_Notam(nid);
				else      keepBtn.Click += (s, e) => Keep_Notam(nid);
				tabPage1.Controls.Add(keepBtn);

				if (kept)
					AddStationChips(tabPage1, notam_ID, Impact, supStored, notam_text,
						NotamRemarkDefault(notam_text, fromDate, tillDate), Remark, storedSupRef, Top, ctrlLeft, ctrlW);

				if (ctrlBox != null) ctrlBox.SendToBack();
				card.SendToBack();
				Top = Top + cardH + 8;
			}
			conn.Close();
			// Auto-save: impact/SUP chips write immediately, remark textfields on Leave.

			// See the comment at the end of Filter_Notams — resetting scroll only up front
			// (before the rebuild) can be silently overridden once AutoScrollMinSize is
			// recalculated for the new content.
			tabPage1.AutoScrollPosition = new Point(0, 0);
		}

		// Stations-tab impact chips — immediate write (no SUBMIT). Assigning an impact also
		// keeps the NOTAM; clearing it leaves the impact blank.
		private void AddStationChips(Control parent, int notam_ID, string Impact, bool supStored, string notamText,
			string remarkDefault, string storedRemark, string storedSupRef, int Top, int ctrlLeft, int ctrlW)
		{
			string[] labels = { "APT CLSD", "CAT I", "No ILS", "Not ALTN", "Fuel", "MISC", "RWY" };
			int pad = 10, areaLeft = ctrlLeft + pad, areaW = ctrlW - 2 * pad;
			int colW = areaW / 4, chipW = colW - 6;
			int[] colX = { areaLeft, areaLeft + colW, areaLeft + 2 * colW, areaLeft + 3 * colW };
			int[] cols = { colX[0], colX[1], colX[2], colX[0], colX[1], colX[2], colX[3] };
			int[] tops = { Top+44, Top+44, Top+44, Top+68, Top+68, Top+68, Top+68 };
			for (int i = 0; i < 7; i++)
			{
				string code = _impactOrder[i];
				CheckBox chk = CreateChip(parent, labels[i], cols[i], tops[i], chipW, Impact == code, code);
				int id = notam_ID; string c = code; CheckBox cb = chk;
				chk.CheckedChanged += (s, e) => StationImpactSet(id, c, cb.Checked);
			}

			// SUP — independent, immediate write
			CheckBox sup = CreateChip(parent, "SUP", colX[3], Top+44, chipW, supStored, "AS");
			int sid = notam_ID; string txt = notamText; CheckBox sb = sup;
			sup.CheckedChanged += (s, e) => StationSupSet(sid, txt, sb.Checked);

			// Remark textfields shown when impact / SUP is set (immediate save on leave).
			// "AS" is the legacy SUP-as-impact code — it is SUP, not an impact remark.
			bool impactOn = HasImpact(Impact) && Impact != "AS";
			if (impactOn || supStored)
			{
				int rowLeft = areaLeft, rowTop = Top + 94, rowW = areaW, gap = 8;
				if (impactOn && supStored)
				{
					AddStationRemark(parent, notam_ID, false, rowLeft, rowTop, rowW * 2 / 3,
						storedRemark != "" ? storedRemark : remarkDefault);
					AddStationRemark(parent, notam_ID, true, rowLeft + rowW * 2 / 3 + gap, rowTop, rowW / 3 - gap, storedSupRef);
				}
				else if (impactOn)
					AddStationRemark(parent, notam_ID, false, rowLeft, rowTop, rowW,
						storedRemark != "" ? storedRemark : remarkDefault);
				else
					AddStationRemark(parent, notam_ID, true, rowLeft, rowTop, rowW, storedSupRef);
			}
		}

		// A remark / SUP-ref textfield on the station view (persisted on Leave AND by SAVE).
		private void AddStationRemark(Control parent, int notam_ID, bool isSup, int left, int top, int width, string text)
		{
			TextBox tb = new TextBox { Tag = "dispose", Left = left, Top = top, Width = width, Height = 24, Text = text };
			int id = notam_ID; bool sup = isSup; TextBox box = tb;
			tb.Leave += (s, e) => StationSaveRemark(id, sup, box.Text);
			parent.Controls.Add(tb);
		}

		// Save remark / SupRef without re-rendering (the field already shows the text).
		void StationSaveRemark(int notam_ID, bool isSup, string text)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand u = new OleDbCommand(
				isSup ? "UPDATE filteredNotams_table SET SupRef=? WHERE ID=?"
				      : "UPDATE filteredNotams_table SET Remark=? WHERE ID=?", conn);
			u.Parameters.AddWithValue("?", text);
			u.Parameters.AddWithValue("?", notam_ID);
			u.ExecuteNonQuery();
			conn.Close();
		}

		// Toggle an impact for a Stations-tab NOTAM and reload.
		void StationImpactSet(int notam_ID, string code, bool on)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand u;
			if (on)
			{
				// New impact -> reset the impact remark so the new impact shows its own default
				u = new OleDbCommand("UPDATE filteredNotams_table SET Impact=?, Remark='', Status='K' WHERE ID=?", conn);
				u.Parameters.AddWithValue("?", code);
				u.Parameters.AddWithValue("?", notam_ID);
			}
			else
			{
				u = new OleDbCommand("UPDATE filteredNotams_table SET Impact='', Remark='' WHERE ID=?", conn);
				u.Parameters.AddWithValue("?", notam_ID);
			}
			u.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
		}

		// Toggle SUP for a Stations-tab NOTAM (independent of impact) and reload.
		void StationSupSet(int notam_ID, string notamText, bool on)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand u;
			if (on)
			{
				u = new OleDbCommand("UPDATE filteredNotams_table SET Sup='Yes', SupRef=?, Status='K' WHERE ID=?", conn);
				u.Parameters.AddWithValue("?", ExtractSupRef(notamText));
				u.Parameters.AddWithValue("?", notam_ID);
			}
			else
			{
				u = new OleDbCommand("UPDATE filteredNotams_table SET Sup='', SupRef='' WHERE ID=?", conn);
				u.Parameters.AddWithValue("?", notam_ID);
			}
			u.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
		}

		void Keep_Notam(int notam_ID)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Status='K' WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			RefreshCurrentView();
		}

		void Ignore_Notam(int notam_ID)
		{
			_autoKeepSkip.Add(notam_ID);   // don't auto-re-keep a NOTAM the dispatcher ignored
			_autoKeptIds.Remove(notam_ID);
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Status='' WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			RefreshCurrentView();
		}

		void Impact_Notam(int notam_ID, string I)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand qry = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE ID=?", conn);
			qry.Parameters.AddWithValue("?", notam_ID);
			OleDbDataReader r = qry.ExecuteReader();
			string Impact = "";
			while (r.Read())
				if (!r.IsDBNull(13)) Impact = r.GetString(13);
			conn.Close();

			conn.Open();
			OleDbCommand upd;
			if (Impact == "")
			{
				upd = new OleDbCommand("UPDATE filteredNotams_table SET Impact=? WHERE ID=?", conn);
				upd.Parameters.AddWithValue("?", I);
				upd.Parameters.AddWithValue("?", notam_ID);
			}
			else
			{
				upd = new OleDbCommand("UPDATE filteredNotams_table SET Impact='' WHERE ID=?", conn);
				upd.Parameters.AddWithValue("?", notam_ID);
			}
			upd.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
			Filter_Notams();
		}

		void Remark_Notam(int notam_ID, string remark)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Remark=? WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", remark);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
		}

		void Btn_submitNotamsClick(object sender, EventArgs e)
		{
			string AP = Lbl_location.Text;
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();

			// Persist the final checkbox/remark state of each kept NOTAM
			foreach (int id in _pendImpactChks.Keys)
			{
				CheckBox[] chks = _pendImpactChks[id];
				string code = "";
				for (int i = 0; i < chks.Length; i++)
					if (chks[i].Checked) { code = _impactOrder[i]; break; }
				bool   supOn  = _pendSupChk.ContainsKey(id) && _pendSupChk[id].Checked;
				string sup    = supOn ? "Yes" : "";
				string remark = (code != "" && _pendRemark.ContainsKey(id)) ? _pendRemark[id].Text : "";
				string supRef = (supOn && _pendSupRemark.ContainsKey(id)) ? _pendSupRemark[id].Text : "";

				OleDbCommand u = new OleDbCommand(
					"UPDATE filteredNotams_table SET Impact=?, Sup=?, Remark=?, SupRef=? WHERE ID=?", conn);
				u.Parameters.AddWithValue("?", code);
				u.Parameters.AddWithValue("?", sup);
				u.Parameters.AddWithValue("?", remark);
				u.Parameters.AddWithValue("?", supRef);
				u.Parameters.AddWithValue("?", id);
				u.ExecuteNonQuery();
			}

			OleDbCommand cmd = new OleDbCommand(
				"UPDATE filteredNotams_table SET Checked='Y' WHERE (location=?) AND (Checked='N')", conn);
			cmd.Parameters.AddWithValue("?", AP);
			cmd.ExecuteNonQuery();
			conn.Close();
			Filter_Notams();
		}

		// Collapse a multi-flag suggestion to a single Impact code by severity priority
		private static string SuggestedSingleCode(ImpactSuggestion s)
		{
			if (s.APClsd)  return "A";
			if (s.NoILS)   return "N";
			if (s.CatI)    return "C";
			if (s.NotAltn) return "D";
			if (s.Fuel)    return "F";
			return "";
		}

		// Filter-tab impact checkboxes: visual-only until SUBMIT. AUTO state = suggested
		// but not yet stored (yellow). Impact group is radio (single choice). SUP independent.
		private void AddFilterCheckboxes(int notam_ID, string storedImpact, bool stored, string sugCode,
			bool supStored, bool supSug, string notamText, string remarkDefault, string storedRemark, string storedSupRef, int Top, int ctrlLeft, int ctrlW)
		{
			_pendRemarkDefault[notam_ID] = remarkDefault;
			string[] labels = { "APT CLSD", "CAT I", "No ILS", "Not ALTN", "Fuel", "MISC", "RWY" };

			// Lay the 4 columns strictly inside the control box [ctrlLeft, ctrlLeft+ctrlW].
			int pad      = 10;
			int areaLeft = ctrlLeft + pad;
			int areaW    = ctrlW - 2 * pad;
			int colW     = areaW / 4;
			int chipW    = colW - 6;
			int[] colX   = { areaLeft, areaLeft + colW, areaLeft + 2 * colW, areaLeft + 3 * colW };
			int[] cols   = { colX[0], colX[1], colX[2], colX[0], colX[1], colX[2], colX[3] };
			int[] tops   = { Top+44, Top+44, Top+44, Top+68, Top+68, Top+68, Top+68 };

			bool supAuto = !supStored && supSug;

			// Both textboxes live inside a fixed-position container panel (within the box).
			Panel remarkRow = new Panel { Tag="dispose", Top=Top+94, Left=areaLeft, Width=areaW, Height=26 };
			tabPage1.Controls.Add(remarkRow);

			// Impact remark textbox: stored remark > impact first-line > empty
			string remarkInit;
			if (stored) remarkInit = storedRemark;
			else if (sugCode != "") remarkInit = remarkDefault;
			else remarkInit = "";
			TextBox remark = new TextBox { Top=0, Left=0, Size=new Size(250,24), Text=remarkInit };
			_pendRemark[notam_ID] = remark;
			remarkRow.Controls.Add(remark);

			CheckBox[] chks = new CheckBox[7];
			for (int i = 0; i < 7; i++)
			{
				string code = _impactOrder[i];
				bool isAuto  = !stored && sugCode == code;
				bool isOn    = stored ? (storedImpact == code) : isAuto;
				int idx = i, nid = notam_ID;

				CheckBox chk = CreateChip(tabPage1, labels[i], cols[i], tops[i], chipW, isOn, code);
				chk.CheckedChanged += (s, ev) => FilterImpactToggled(nid, idx, notamText);
				chks[i] = chk;
			}
			_pendImpactChks[notam_ID] = chks;

			// SUP — fully independent of the impact state
			CheckBox sup = CreateChip(tabPage1, "SUP", colX[3], Top+44, chipW, supStored || supSug, "AS");
			int snid = notam_ID;
			sup.CheckedChanged += (s, ev) => FilterSupToggled(snid, notamText);
			_pendSupChk[notam_ID] = sup;

			// Independent SUP reference textbox (separate from the impact remark)
			string supInit;
			if (supStored) supInit = storedSupRef;
			else if (supAuto) supInit = ExtractSupRef(notamText);
			else supInit = "";
			TextBox supRef = new TextBox { Top=0, Left=0, Size=new Size(220,24), Text=supInit };
			_pendSupRemark[notam_ID] = supRef;
			remarkRow.Controls.Add(supRef);

			LayoutRemarkBoxes(notam_ID);   // initial visibility/width
		}

		// Show only the relevant textbox(es), sized to fit the panel width (panel-relative
		// coords): one box -> full width; two boxes -> 2/3 impact remark, 1/3 SUP reference.
		private void LayoutRemarkBoxes(int notam_ID)
		{
			if (!_pendRemark.ContainsKey(notam_ID) || !_pendSupRemark.ContainsKey(notam_ID)) return;
			const int GAP = 8;
			TextBox r  = _pendRemark[notam_ID];
			TextBox sp = _pendSupRemark[notam_ID];
			int W = (r.Parent != null) ? r.Parent.Width : 770;

			bool impactOn = false;
			if (_pendImpactChks.ContainsKey(notam_ID))
				foreach (CheckBox c in _pendImpactChks[notam_ID]) if (c.Checked) { impactOn = true; break; }
			bool supOn = _pendSupChk.ContainsKey(notam_ID) && _pendSupChk[notam_ID].Checked;

			if (impactOn && supOn)
			{
				r.Visible = true;  r.Left = 0;               r.Width = W * 2 / 3;
				sp.Visible = true; sp.Left = W * 2 / 3 + GAP; sp.Width = W / 3 - GAP;
			}
			else if (impactOn) { r.Visible = true; r.Left = 0; r.Width = W; sp.Visible = false; }
			else if (supOn)    { sp.Visible = true; sp.Left = 0; sp.Width = W; r.Visible = false; }
			else               { r.Visible = false; sp.Visible = false; }
		}

		// SUP selection pre-fills the independent SUP textbox with the SUP reference if empty
		private void FilterSupToggled(int notam_ID, string notamText)
		{
			if (_pendSupChk.ContainsKey(notam_ID) && _pendSupChk[notam_ID].Checked &&
			    _pendSupRemark.ContainsKey(notam_ID) && _pendSupRemark[notam_ID].Text.Trim() == "")
			{
				string r = ExtractSupRef(notamText);
				if (r != "") _pendSupRemark[notam_ID].Text = r;
			}
			if (_pendSupChk.ContainsKey(notam_ID)) StyleChk(_pendSupChk[notam_ID], "AS");
			LayoutRemarkBoxes(notam_ID);
		}

		// Radio behaviour within the impact group + remark auto-fill
		private void FilterImpactToggled(int notam_ID, int idx, string notamText)
		{
			if (!_pendImpactChks.ContainsKey(notam_ID)) return;
			CheckBox[] chks = _pendImpactChks[notam_ID];
			if (chks[idx].Checked)
			{
				for (int i = 0; i < chks.Length; i++)
					if (i != idx && chks[i].Checked) chks[i].Checked = false;

				// Pre-fill remark with first line of NOTAM text if empty
				if (_pendRemark.ContainsKey(notam_ID) && _pendRemark[notam_ID].Text.Trim() == "" &&
				    _pendRemarkDefault.ContainsKey(notam_ID))
				{
					_pendRemark[notam_ID].Text = _pendRemarkDefault[notam_ID];
				}
			}
			for (int i = 0; i < chks.Length; i++) StyleChk(chks[i], _impactOrder[i]);
			LayoutRemarkBoxes(notam_ID);
		}


		void Btn_ICAOClick(object sender, EventArgs e)                        { ICAO_Notams(); }
		void Btn_filterNewClick(object sender, EventArgs e)                   { Filter_Notams(); }
		void ChckBox_SeeIgnoredCheckedChanged(object sender, EventArgs e)     { ICAO_Notams(); }

		// Quick-access duplicate of the "DB Update" button (tabPage3), reachable without
		// leaving the NOTAM Filter tab. Runs the same download/dedupe/sync flow.
		void Btn_dbUpdateQuickClick(object sender, EventArgs e)
		{
			RunDbUpdatePipeline(delegate { RefreshCurrentView(); });
		}

		// Shows when the "DB Update" pipeline (download + dedupe + sync) was last actually
		// run to completion — not the local .mdb file's timestamp, which also changes on
		// V: drive sync (StartApp/EndApp) and would misleadingly look "fresh" even if the
		// dispatcher never clicked DB Update. The timestamp is written by
		// SaveLastDbUpdateTimestamp() at the end of a successful RunDbUpdatePipeline.
		void RefreshLastDbUpdateLabel()
		{
			try
			{
				string tsPath = System.IO.Path.Combine(Application.StartupPath, "last_db_update.txt");
				DateTime ts;
				Lbl_lastDbUpdate.Text = System.IO.File.Exists(tsPath) &&
					DateTime.TryParse(System.IO.File.ReadAllText(tsPath), System.Globalization.CultureInfo.InvariantCulture,
						System.Globalization.DateTimeStyles.None, out ts)
					? "MAJ: " + ts.ToString("dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture)
					: "MAJ: n/a";
			}
			catch { Lbl_lastDbUpdate.Text = "MAJ: n/a"; }
		}

		// Repositions the top-bar controls (ICAO box excluded — no, included) with raw pixel
		// values matching the dynamically-created card column (Left=505). Designer controls
		// are subject to WinForms DPI auto-scaling while the code-built cards are not, so on
		// a non-design-time DPI the two families drift apart; setting Left/Top here in code
		// keeps them pixel-identical to the card column regardless of DPI.
		void AlignTopBar()
		{
			Btn_filterNew.Left = 7;        Btn_filterNew.Top = 8;
			Btn_dbUpdateQuick.Left = 205;   Btn_dbUpdateQuick.Top = 8;
			Lbl_lastDbUpdate.Left = Btn_dbUpdateQuick.Right + 4;   Lbl_lastDbUpdate.Top = 16;
			Lbl_ICAO.Left = 505;            Lbl_ICAO.Top = 13;
			TxtBox_ICAO.Left = 575;         TxtBox_ICAO.Top = 9;
			Btn_ICAO.Left = 675;            Btn_ICAO.Top = 8;
		}

		// Re-render whichever mode is currently shown on tabPage1
		void RefreshCurrentView()
		{
			if (_stationMode) ICAO_Notams();
			else              Filter_Notams();
		}

		// ── private helpers ──────────────────────────────────────────────────

		private string FormatDate(string raw)
		{
			if (raw.Length < 16) return raw;
			string d = raw.Substring(0, 16);
			return d.Substring(8,2) + " " + MonthAbbrev(d.Substring(5,2)) + " " + d.Substring(0,4) + " " + d.Substring(11,5) + "z";
		}

		private void AddNotamLabel(Control parent, FontFamily font, string text, int top, int left, int width,
			Color color, bool bold = false, float size = 10f)
		{
			Label lbl = new Label
			{
				Font      = new Font(font, size, bold ? FontStyle.Bold : FontStyle.Regular),
				Tag       = "dispose",
				Top       = top, Left = left,
				Size      = new Size(width, 18),
				ForeColor = color, Text = text
			};
			parent.Controls.Add(lbl);
		}

		private static string HighlightKeywordsHtml(string text)
		{
			string s = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
			foreach (string kw in _notamKeywords)
			{
				int idx = 0;
				while (true)
				{
					idx = s.IndexOf(kw, idx, StringComparison.OrdinalIgnoreCase);
					if (idx < 0) break;
					if (!IsWholeWord(s, idx, kw.Length)) { idx += kw.Length; continue; }
					string orig = s.Substring(idx, kw.Length);
					string rep  = "<b style=\"color:#b00000\">" + orig + "</b>";
					s    = s.Substring(0, idx) + rep + s.Substring(idx + kw.Length);
					idx += rep.Length;
				}
			}
			return s.Replace("\n", "<br>");
		}

		private static WebBrowser MakeNotamWebBrowser(string text, int top, int left, int height, bool kept)
		{
			string bg = kept ? "#dce3ff" : "#ffffff";
			WebBrowser wb = new WebBrowser
			{
				Tag = "dispose", Top = top, Left = left, Size = new Size(550, height),
				ScrollBarsEnabled = false
			};
			wb.DocumentText =
				"<html><head><style>" +
				"body{margin:2px;padding:4px 8px;background:" + bg + ";font-family:'Courier New',monospace;font-size:12px;overflow:hidden;line-height:1.6;border:1px solid #222}" +
				"</style></head><body>" +
				HighlightKeywordsHtml(text) +
				"</body></html>";
			return wb;
		}

		private Button MakeKeepButton(FontFamily font, string status, Point location)
		{
			Button b = new Button
			{
				Tag       = "dispose", Location = location,
				Size      = status == "K" ? new Size(58,24) : new Size(48,24),
				Text      = status == "K" ? "Ignore" : "Keep",
				ForeColor = Color.White,
				BackColor = status == "K" ? Color.FromArgb(84,110,122) : Color.FromArgb(100,149,237),
				Font      = new Font("Segoe UI", 8f),
				FlatStyle = FlatStyle.Flat
			};
			b.FlatAppearance.BorderSize = 0;
			return b;
		}

		private bool HasImpact(string impact)
		{
			return impact == "A" || impact == "C" || impact == "N" || impact == "D" ||
			       impact == "F" || impact == "M" || impact == "AS" || impact == "R";
		}

		private void AddImpactCheckboxes(Control parent, int notam_ID, string Impact, int Top,
			int col1, int col2, int col3, int col4,
			Dictionary<int,CheckBox> CLSD,   Dictionary<int,CheckBox> CATI,
			Dictionary<int,CheckBox> NILS,   Dictionary<int,CheckBox> NOALTN,
			Dictionary<int,CheckBox> FUEL,   Dictionary<int,CheckBox> MISC,
			Dictionary<int,CheckBox> AIPSUP, Dictionary<int,CheckBox> RWYCLSD)
		{
			int nid = notam_ID;
			CLSD[notam_ID]    = AddImpactChk(parent, "APT CLSD", Top+44, col1, Impact=="A",  (s,e) => Impact_Notam(nid,"A"));
			CATI[notam_ID]    = AddImpactChk(parent, "APT CATI", Top+44, col2, Impact=="C",  (s,e) => Impact_Notam(nid,"C"));
			NILS[notam_ID]    = AddImpactChk(parent, "No ILS",   Top+44, col3, Impact=="N",  (s,e) => Impact_Notam(nid,"N"));
			NOALTN[notam_ID]  = AddImpactChk(parent, "Not ALTN", Top+68, col1, Impact=="D",  (s,e) => Impact_Notam(nid,"D"));
			FUEL[notam_ID]    = AddImpactChk(parent, "Fuel",     Top+68, col2, Impact=="F",  (s,e) => Impact_Notam(nid,"F"));
			MISC[notam_ID]    = AddImpactChk(parent, "MISC",     Top+68, col3, Impact=="M",  (s,e) => Impact_Notam(nid,"M"));
			AIPSUP[notam_ID]  = AddImpactChk(parent, "SUP",      Top+44, col4, Impact=="AS", (s,e) => Impact_Notam(nid,"AS"));
			RWYCLSD[notam_ID] = AddImpactChk(parent, "RWY",      Top+68, col4, Impact=="R",  (s,e) => Impact_Notam(nid,"R"));
		}

		private CheckBox AddImpactChk(Control parent, string text, int top, int left, bool chked, EventHandler handler)
		{
			CheckBox chk = new CheckBox { Tag="dispose", Top=top, Left=left, Text=text, Size=new Size(80,25), Checked=chked };
			chk.CheckedChanged += handler;
			parent.Controls.Add(chk);
			return chk;
		}
	}
}
