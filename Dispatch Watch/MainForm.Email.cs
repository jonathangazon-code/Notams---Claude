using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private const string ReportsDir =
			@"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";

		// Creates the EmailRecipients table if missing (idempotent).
		public void EnsureEmailTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE EmailRecipients ([Email] TEXT(255))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		List<string> LoadRecipients()
		{
			List<string> list = new List<string>();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Email FROM EmailRecipients ORDER BY Email", conn).ExecuteReader();
				while (r.Read())
					if (!r.IsDBNull(0) && r.GetString(0).Trim() != "") list.Add(r.GetString(0).Trim());
				conn.Close();
			}
			catch { }
			return list;
		}

		// Returns the user's default Outlook signature as HTML (read from the signature files),
		// or "" if none. Avoids GetInspector, which breaks programmatic .Send.
		//
		// inlineImages: any image the signature references (e.g. a company logo) gets rewritten
		// to a "cid:..." reference and its real file registered in this dictionary, the same
		// (cid -> temp/real path) map the runway-diagram images already use — the caller attaches
		// every entry as a hidden file whose PR_ATTACH_CONTENT_ID matches the cid (see
		// Btn_sendReportsClick/Btn_sendTafReportClick's own "InlineImages" step). A "file:///..."
		// reference — what this used to rewrite to — only ever resolves on the SENDER's own
		// machine (it points at a path under this user's local AppData); every recipient just
		// sees a broken-image placeholder, since they have no access to that path at all.
		string ReadDefaultSignatureHtml(Dictionary<string, string> inlineImages)
		{
			try
			{
				string sigDir = Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Signatures");
				if (!Directory.Exists(sigDir)) return "";

				// Default "new message" signature name, stored under Outlook's MailSettings.
				string name = "";
				foreach (string ver in new string[] { "16.0", "15.0", "14.0" })
				{
					try
					{
						RegistryKey key = Registry.CurrentUser.OpenSubKey(
							@"Software\Microsoft\Office\" + ver + @"\Common\MailSettings");
						if (key != null)
						{
							object v = key.GetValue("NewSignature");
							if (v is byte[]) name = Encoding.Unicode.GetString((byte[])v).TrimEnd('\0');
							else if (v != null) name = v.ToString();
							if (name != "") break;
						}
					}
					catch { }
				}

				string file = "";
				if (name != "")
				{
					string cand = Path.Combine(sigDir, name + ".htm");
					if (File.Exists(cand)) file = cand;
				}
				if (file == "")
				{
					string[] htms = Directory.GetFiles(sigDir, "*.htm");
					if (htms.Length > 0) file = htms[0];
				}
				if (file == "") return "";

				string html = File.ReadAllText(file);

				// Outlook signatures store their images in a sibling "<name>_files" folder,
				// referenced by relative "<name>_files/imageNNN.png" src attributes. Each one that
				// actually exists on disk gets replaced with a "cid:sig_..." reference, and its
				// real path registered for the caller to attach as a hidden image.
				string imgDir = Path.GetFileNameWithoutExtension(file) + "_files";
				string imgDirPath = Path.Combine(sigDir, imgDir);
				if (inlineImages != null && Directory.Exists(imgDirPath))
				{
					html = Regex.Replace(html, "\"" + Regex.Escape(imgDir) + "/([^\"]+)\"", delegate(Match m)
					{
						string fileName = m.Groups[1].Value;
						string fullPath = Path.Combine(imgDirPath, fileName);
						if (!File.Exists(fullPath)) return m.Value;
						string cid = "sig_" + Regex.Replace(fileName, @"[^A-Za-z0-9]", "_");
						inlineImages[cid] = fullPath;
						return "\"cid:" + cid + "\"";
					});
				}

				return "<br>" + html;
			}
			catch { return ""; }
		}

		// Sends today's two PDF reports via Outlook (late-bound COM, direct send).
		void Btn_sendReportsClick(object sender, EventArgs e)
		{
			if (!EnsureWriterOrWarn()) return;
			string today    = DateTime.Now.ToString("yyMMdd");
			string notamPdf = Path.Combine(ReportsDir, today + "-NOTAMS_report.pdf");
			string supPdf   = Path.Combine(ReportsDir, today + "-AIP_SUP_report.pdf");
			string fsPdf    = Path.Combine(ReportsDir, today + "-Flight_Schedule.pdf");

			List<string> missing = new List<string>();
			if (!File.Exists(notamPdf)) missing.Add(Path.GetFileName(notamPdf));
			if (!File.Exists(supPdf))   missing.Add(Path.GetFileName(supPdf));
			if (!File.Exists(fsPdf))    missing.Add(Path.GetFileName(fsPdf));
			if (missing.Count > 0)
			{
				MessageBox.Show("Today's PDF report(s) not found:\n" + string.Join("\n", missing.ToArray()) +
					"\n\nExport all three reports to PDF first.", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			List<string> rcp = LoadRecipients();
			if (rcp.Count == 0)
			{
				MessageBox.Show("No recipients defined. Add at least one address.", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			string to = string.Join(";", rcp.ToArray());

			if (MessageBox.Show("Send today's reports to " + rcp.Count + " recipient(s)?",
				"Send Reports", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

			string titleDate = DateTime.Now.ToString("ddMMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToUpper(); // e.g. 09JUL2026
			string subject   = "Ops Briefing – NOTAMs, AIP SUP & Flight Schedule - " + titleDate;

			// The email body is the Conflict tab's own report (MainForm.Conflict.cs), so the
			// two never drift apart — same intro sentence, same impact sections/colors/badges,
			// just prefixed with a note about the three attached PDFs. Runway diagrams are
			// rasterized (rasterDiagrams=true, unlike the live tab's VML) and embedded as
			// cid: references (cidImages=true) rather than data-URIs — Outlook's Word engine
			// doesn't render data-URI images in an HTML body, only CID-attached ones.
			Dictionary<string, string> inlineImages;
			string bodyHtml = BuildConflictReportHtml(true, true, out inlineImages);
			// .attachNote is defined in the shared <style> block BuildConflictReportHtml just
			// built (bigger, green-accented, same treatment as the report's other explanatory
			// text) — this paragraph gets inserted into that same HTML document below, so the
			// class is available here without repeating the styling inline. Each report name is
			// also a real, underlined file:// link straight to its V: path (same as the reports
			// already attached) — since Reports lives on the shared drive, any recipient with
			// access to V: can open the live file directly, not just the attached snapshot.
			string notamLink = "<a href=\"" + new Uri(notamPdf).AbsoluteUri + "\" style=\"color:inherit;text-decoration:underline\">the complete NOTAMs analysis (Impact summary + complete Network)</a>";
			string supLink   = "<a href=\"" + new Uri(supPdf).AbsoluteUri   + "\" style=\"color:inherit;text-decoration:underline\">the list of AIP SUPs loaded in Aviobook</a>";
			string fsLink    = "<a href=\"" + new Uri(fsPdf).AbsoluteUri    + "\" style=\"color:inherit;text-decoration:underline\">the Flight Schedule used for the coming 7 days</a>";
			string attachmentNote =
				"<p class=\"attachNote\">" +
				"In attachment you will find:<br>" + notamLink + "<br>" + supLink + "<br>" + fsLink + "</p>";

			string step = "init";
			try
			{
				Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
				if (outlookType == null)
				{
					MessageBox.Show("Outlook is not installed or registered on this machine.", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				step = "CreateInstance";
				object outlook = Activator.CreateInstance(outlookType);
				step = "CreateItem";
				object mail = outlookType.InvokeMember("CreateItem", BindingFlags.InvokeMethod, null, outlook, new object[] { 0 }); // 0 = olMailItem
				Type mt = mail.GetType();
				step = "Subject";
				mt.InvokeMember("Subject", BindingFlags.SetProperty, null, mail, new object[] { subject });

				// Recipients: add each, resolve, and report any that Outlook cannot recognise.
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
						"\n\nFix or remove them in the recipients list, then try again.",
						"Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				// Body (the Conflict report + attachment note) + default signature read from
				// the Outlook signature files (no GetInspector, which conflicts with .Send) —
				// both inserted just before </body> so they land inside the Conflict report's
				// own <html> document instead of dangling after its closing tag.
				step = "Body";
				string fullBody = bodyHtml;
				int bodyCloseIdx = fullBody.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
				string tail = attachmentNote + ReadDefaultSignatureHtml(inlineImages);
				fullBody = bodyCloseIdx >= 0 ? fullBody.Insert(bodyCloseIdx, tail) : fullBody + tail;
				mt.InvokeMember("HTMLBody", BindingFlags.SetProperty, null, mail, new object[] { fullBody });

				step = "Attachments";
				object atts = mt.InvokeMember("Attachments", BindingFlags.GetProperty, null, mail, null);
				Type at = atts.GetType();
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { notamPdf });
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { fsPdf });
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { supPdf });

				// Runway-diagram PNGs: attached as hidden files whose PR_ATTACH_CONTENT_ID
				// matches the "cid:..." reference BuildConflictReportHtml embedded in the body
				// — the standard way to inline an image in an HTML email via Outlook COM.
				step = "InlineImages";
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

				MessageBox.Show("Reports sent to " + rcp.Count + " recipient(s).", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				string msg = (ex.InnerException != null) ? ex.InnerException.Message : ex.Message;
				MessageBox.Show("Failed to send via Outlook (step: " + step + "):\n" + msg, "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// inlineImages now also holds the dispatcher's REAL Outlook signature image
				// file(s) (ReadDefaultSignatureHtml), not just the runway-diagram PNGs this
				// cleanup was originally written for — only delete what's actually a temp copy
				// (under Path.GetTempPath()), or a resend would permanently destroy their
				// signature's logo.
				string tempRoot = Path.GetTempPath();
				foreach (string p in inlineImages.Values)
					if (p.StartsWith(tempRoot, StringComparison.OrdinalIgnoreCase))
						try { if (File.Exists(p)) File.Delete(p); } catch { }
			}
		}
	}
}
