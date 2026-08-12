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
			try
			{
				string ico = System.IO.Path.Combine(Application.StartupPath, "icon.ico");
				if (System.IO.File.Exists(ico)) this.Icon = new System.Drawing.Icon(ico);
			}
			catch { /* icon optional */ }

			// One-time OCC.mdb -> ICAO_storedNotams.mdb migration, run against the shared V:
			// files directly (not wherever this process happens to be running from) so it
			// takes effect exactly once regardless of who launches first — must run before
			// EnsureLocalInstall(), which may relocate this process off V: and never return.
			EnsureOccMigrated();

			// Deployment: install-to-local-folder/auto-update may relaunch and Environment.Exit
			// before returning, so it runs before anything else touches the DB. User identity
			// and the Writer/Reader lock both need to be resolved before StartApp()'s download —
			// see MainForm.Deployment.cs.
			EnsureLocalInstall();
			_currentUserEmail = GetCurrentUserEmail();
			EnsureUsersTable();
			UpsertCurrentUser(_currentUserEmail);
			AcquireOrPromptForLock();

			StartApp();
			EnsureSchema();
			EnsureKeywordsTable();
			EnsureEmailTable();
			EnsureRunwaysTable();
			PreloadCsvGeoAsync();   // load runways.csv geo index off the UI thread
			EnsureAirportNameColumn();
			PreloadAirportNamesAsync();   // load airports.csv name index off the UI thread
			EnsureArchiveConfig();
			EnsureFlightScheduleTable();
			EnsureFlightSchedFolder();
			RefreshLastDbUpdateLabel();
			LoadKeywords();
			LoadStationsCache();
			tab_RWYs();
			Airport_List();
			Keywords_Refresh();
			Recipients_Refresh();
			if (!IsWriter) ApplyReaderModeUi();
			StartIdleAutoSave();
			this.FormClosing += MainForm_FormClosing;
			// Run the first Filter render once the window is shown (and maximized) so the
			// layout uses the real tab width rather than the smaller design-time size.
			this.Shown += delegate { _formShown = true; Filter_Notams(); };

			// The NOTAM Filter/Station cards compute their width from tabPage1.ClientSize.Width
			// at render time and don't auto-resize afterward (they're absolute-positioned, not
			// docked/anchored) — without this, widening the window past its initial maximized
			// size leaves the right column stuck at its old, narrower width. Debounced via a
			// short timer so a drag-resize doesn't re-render (and re-query the DB) on every
			// intermediate frame — only once, shortly after the size settles.
			_resizeDebounce = new Timer { Interval = 200 };
			_resizeDebounce.Tick += delegate { _resizeDebounce.Stop(); if (_formShown) RefreshCurrentView(); };
			this.Resize += delegate { _resizeDebounce.Stop(); _resizeDebounce.Start(); };
		}

		private bool _formShown;
		private Timer _resizeDebounce;

		// Ctrl+S on the NOTAM Filter tab == clicking SUBMIT. ProcessCmdKey (not a control's own
		// KeyDown) catches the shortcut regardless of which child control currently has focus —
		// the NOTAM cards/chips/textboxes on this tab are all dynamically built, so there's no
		// single control to attach a KeyDown handler to. Only fires in the automatic Filter New
		// NOTAMS flow (Btn_submitNotams.Visible is false in the manual ICAO station view, which
		// has no Submit step at all — every edit there already saves immediately).
		protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
		{
			if (keyData == (Keys.Control | Keys.S) && tabControl1.SelectedTab == tabPage1 && Btn_submitNotams.Visible)
			{
				Btn_submitNotamsClick(this, EventArgs.Empty);
				return true;
			}
			return base.ProcessCmdKey(ref msg, keyData);
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (_lockHeartbeatTimer != null) _lockHeartbeatTimer.Stop();
			if (_idleCheckTimer != null) _idleCheckTimer.Stop();
			ReleaseLock();   // before the upload below, so the next launcher never sees a stale "in use" state
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
					LogDeploy("StartApp: source not found: " + sourcePath);
					MessageBox.Show("Fichier source introuvable :\n" + sourcePath, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				bool destExists = File.Exists(destPath);
				DateTime destTime = destExists ? File.GetLastWriteTime(destPath) : DateTime.MinValue;
				DateTime sourceTime = File.GetLastWriteTime(sourcePath);
				LogDeploy("StartApp: destExists=" + destExists + " destTime=" + destTime.ToString("o") + " sourceTime=" + sourceTime.ToString("o"));
				if (destExists && destTime >= sourceTime)
				{
					LogDeploy("StartApp: local DB considered up to date — no download.");
					return;
				}
				File.Copy(sourcePath, destPath, true);
				LogDeploy("StartApp: downloaded DB to " + destPath + " (new local write time=" + File.GetLastWriteTime(destPath).ToString("o") + ")");
			}
			catch (Exception ex)
			{
				LogDeploy("StartApp: EXCEPTION: " + ex);
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
		void Btn_delWithdrawnedClick(object sender, EventArgs e)  { deleteWithdrawnedNotams(); }
		void Btn_printReportClick(object sender, EventArgs e)        { Web_report.Print(); }
		void Btn_exportReportClick(object sender, EventArgs e)       { ExportNotamReportPdf(); }
		void Btn_Sup_printReportClick(object sender, EventArgs e)    { Web_Sup_report.Print(); }
		void Btn_Sup_exportReportClick(object sender, EventArgs e)   { ExportToPdf(Web_Sup_report, "AIP_SUP_report"); }
		void Btn_fsExportReportClick(object sender, EventArgs e)     { ExportToPdf(BuildFlightScheduleHtml(), "Flight_Schedule"); }
		void Btn_restartAppClick(object sender, EventArgs e)      { Application.Restart(); }
		void Btn_XMLClick(object sender, EventArgs e)             { GetXML(); }
		void Btn_reloadClick(object sender, EventArgs e)          { Reload_text(); }
	}
}
