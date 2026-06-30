using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

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

			string ddmmyyyy = DateTime.Now.ToString("ddMMyyyy");
			string subject  = "NOTAMs Report and AIP Sup List - " + ddmmyyyy;
			string body     = "Dear all,\r\n" +
				"please find attached the NOTAMs Report and AIP SUP List for today " + ddmmyyyy + " \r\n" +
				"Kind regards,";

			try
			{
				Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
				if (outlookType == null)
				{
					MessageBox.Show("Outlook is not installed or registered on this machine.", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				object outlook = Activator.CreateInstance(outlookType);
				object mail = outlookType.InvokeMember("CreateItem", BindingFlags.InvokeMethod, null, outlook, new object[] { 0 }); // 0 = olMailItem
				Type mt = mail.GetType();
				mt.InvokeMember("To",      BindingFlags.SetProperty, null, mail, new object[] { to });
				mt.InvokeMember("Subject", BindingFlags.SetProperty, null, mail, new object[] { subject });
				mt.InvokeMember("Body",    BindingFlags.SetProperty, null, mail, new object[] { body });

				object atts = mt.InvokeMember("Attachments", BindingFlags.GetProperty, null, mail, null);
				Type at = atts.GetType();
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { notamPdf });
				at.InvokeMember("Add", BindingFlags.InvokeMethod, null, atts, new object[] { supPdf });

				mt.InvokeMember("Send", BindingFlags.InvokeMethod, null, mail, null);

				MessageBox.Show("Reports sent to " + rcp.Count + " recipient(s).", "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to send via Outlook:\n" + ex.Message, "Send Reports", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
