using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Text;
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

		void Recipients_Refresh()
		{
			Lst_Recipients.Items.Clear();
			foreach (string a in LoadRecipients()) Lst_Recipients.Items.Add(a);
		}

		void Btn_addRecipientClick(object sender, EventArgs e)
		{
			string a = TxtBox_recipient.Text.Trim();
			if (a == "" || !a.Contains("@")) return;
			foreach (object it in Lst_Recipients.Items)
				if (string.Equals(it.ToString(), a, StringComparison.OrdinalIgnoreCase)) { TxtBox_recipient.Clear(); return; }

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand ins = new OleDbCommand("INSERT INTO EmailRecipients ([Email]) VALUES (?)", conn);
			ins.Parameters.AddWithValue("?", a);
			ins.ExecuteNonQuery();
			conn.Close();

			TxtBox_recipient.Clear();
			Recipients_Refresh();
		}

		void TxtBox_recipientKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) { Btn_addRecipientClick(sender, e); e.SuppressKeyPress = true; }
		}

		void Btn_removeRecipientClick(object sender, EventArgs e)
		{
			if (Lst_Recipients.SelectedItem == null) return;
			string a = Lst_Recipients.SelectedItem.ToString();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM EmailRecipients WHERE Email=?", conn);
			del.Parameters.AddWithValue("?", a);
			del.ExecuteNonQuery();
			conn.Close();

			Recipients_Refresh();
		}

		// Returns the user's default Outlook signature as HTML (read from the signature files),
		// or "" if none. Avoids GetInspector, which breaks programmatic .Send.
		string ReadDefaultSignatureHtml()
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
				// Point relative image links to the absolute signature folder so they may resolve.
				string imgDir = Path.GetFileNameWithoutExtension(file) + "_files";
				html = html.Replace("\"" + imgDir + "/", "\"file:///" + Path.Combine(sigDir, imgDir).Replace("\\", "/") + "/");
				return "<br>" + html;
			}
			catch { return ""; }
		}

		// Sends today's two PDF reports via Outlook (late-bound COM, direct send).
		void Btn_sendReportsClick(object sender, EventArgs e)
		{
			string today    = DateTime.Now.ToString("yyMMdd");
			string notamPdf = Path.Combine(ReportsDir, today + "-NOTAMS_report.pdf");
			string supPdf   = Path.Combine(ReportsDir, today + "-AIP_SUP_report.pdf");

			List<string> missing = new List<string>();
			if (!File.Exists(notamPdf)) missing.Add(Path.GetFileName(notamPdf));
			if (!File.Exists(supPdf))   missing.Add(Path.GetFileName(supPdf));
			if (missing.Count > 0)
			{
				MessageBox.Show("Today's PDF report(s) not found:\n" + string.Join("\n", missing.ToArray()) +
					"\n\nExport both reports to PDF first.", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
			string subject   = "NOTAMs Report and AIP Sup List - " + titleDate;
			string bodyHtml  = "Dear all,<br><br>" +
				"Please find attached the NOTAMs Report and AIP SUP List for today " + titleDate + " <br><br>" +
				"Kind regards,<br>";

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

				// Body + default signature read from the Outlook signature files (no GetInspector,
				// which conflicts with .Send).
				step = "Body";
				mt.InvokeMember("HTMLBody", BindingFlags.SetProperty, null, mail, new object[] { bodyHtml + ReadDefaultSignatureHtml() });

				step = "Attachments";
				object atts = mt.InvokeMember("Attachments", BindingFlags.GetProperty, null, mail, null);
				Type at = atts.GetType();
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { notamPdf });
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { supPdf });

				step = "Send";
				mt.InvokeMember("Send", BindingFlags.InvokeMethod, null, mail, null);

				MessageBox.Show("Reports sent to " + rcp.Count + " recipient(s).", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				string msg = (ex.InnerException != null) ? ex.InnerException.Message : ex.Message;
				MessageBox.Show("Failed to send via Outlook (step: " + step + "):\n" + msg, "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
