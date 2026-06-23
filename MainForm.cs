/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 24-02-21
 */
using System;
using System.IO;
using System.Windows.Forms;

namespace ICAO_CSV
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
			StartApp();
			LoadStationsCache();
			tab_RWYs();
			Airport_List();
			this.FormClosing += MainForm_FormClosing;
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			EndApp();
		}

		public void StartApp()
		{
			string sourcePath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\ICAO_storedNotams.mdb";
			string destPath   = System.IO.Path.Combine(Application.StartupPath, "ICAO_storedNotams.mdb");

			try
			{
				if (!File.Exists(sourcePath))
				{
					MessageBox.Show("Fichier source introuvable :\n" + sourcePath, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				if (File.Exists(destPath) && File.GetLastWriteTime(destPath) >= File.GetLastWriteTime(sourcePath))
				{
					MessageBox.Show("Local DB already up to date.\n\nNo download needed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				File.Copy(sourcePath, destPath, true);
				MessageBox.Show("DB Downloaded\n\nSaved to:\n" + destPath, "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Erreur lors de la copie :\n" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public void EndApp()
		{
			string localPath  = System.IO.Path.Combine(Application.StartupPath, "ICAO_storedNotams.mdb");
			string vDrivePath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\ICAO_storedNotams.mdb";

			try
			{
				if (!File.Exists(localPath))
				{
					MessageBox.Show("Local database not found:\n" + localPath + "\n\nUpload to V: cancelled.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (!File.Exists(vDrivePath))
				{
					File.Copy(localPath, vDrivePath, true);
					MessageBox.Show("DB uploaded to V: (file created).\n\n" + vDrivePath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (File.GetLastWriteTime(localPath) > File.GetLastWriteTime(vDrivePath))
				{
					File.Copy(localPath, vDrivePath, true);
					MessageBox.Show("DB on V: updated (local version was newer).", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					MessageBox.Show("No upload needed.\nV: is up to date.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error uploading DB to V:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// ── Button event handlers ────────────────────────────────────────────

		void Btn_CSVClick(object sender, EventArgs e)             { GetCSV(); }
		void Btn_splitClick(object sender, EventArgs e)           { Split(); }
		void Btn_newNotamsClick(object sender, EventArgs e)       { NewNotams(); }
		void Btn_delOldClick(object sender, EventArgs e)          { DelOld(); }
		void Btn_analyzeNotamsClick(object sender, EventArgs e)   { Filter_Notams(); }
		void Btn_delWithdrawnedClick(object sender, EventArgs e)  { deleteWithdrawnedNotams(); }
		void Btn_reportClick(object sender, EventArgs e)          { Report(); }
		void Btn_printReportClick(object sender, EventArgs e)        { Web_report.Print(); }
		void Btn_exportReportClick(object sender, EventArgs e)       { ExportToPdf(Web_report, "NOTAMS_report"); }
		void Btn_AIP_Sup_reportClick(object sender, EventArgs e)     { Sup_Report(); }
		void Btn_Sup_printReportClick(object sender, EventArgs e)    { Web_Sup_report.Print(); }
		void Btn_Sup_exportReportClick(object sender, EventArgs e)   { ExportToPdf(Web_Sup_report, "AIP_SUP_report"); }
		void Btn_restartAppClick(object sender, EventArgs e)      { Application.Restart(); }
		void Btn_XMLClick(object sender, EventArgs e)             { GetXML(); }
		void Btn_reloadClick(object sender, EventArgs e)          { Reload_text(); }
	}
}
