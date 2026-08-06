using System;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// ── Local install + auto-update (A/B) ──────────────────────────────
		// The V: app folder is the same one StartApp/EndApp already sync the DB against
		// (MainForm.cs) — kept as a separate constant here since this concerns the app
		// binaries, not the DB file specifically.
		private const string VAppFolder = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP";
		private const string AppExeName = "Dispatch Watch.exe";

		// Files that are per-machine/per-session state, not app binaries — never copied by
		// the install/update flow (a stale local ArchiveConfig.xml or lock file would be
		// actively wrong to overwrite, and .mdb has its own StartApp/EndApp sync). Everything
		// now lives in the single ICAO_storedNotams.mdb (see EnsureOccMigrated below) — there
		// used to be a second OCC.mdb with no sync of its own at all, which is what caused a
		// fresh local install to crash on startup ("Could not find file ...\OCC.mdb"); merging
		// the two removed that whole class of problem rather than patching around it.
		private static readonly string[] _deployExcludeExtensions = { ".mdb", ".lock" };
		private static readonly string[] _deployExcludeFiles =
			{ "ArchiveConfig.xml", "last_db_update.txt", "MmScheduleCursor.txt", "wkhtmltopdf_log.txt", "_export_temp.html" };

		// Runs once per startup, before anything else touches the DB. Two things it can do:
		// (1) first run for this dispatcher (or launched straight off V:, e.g. via the
		// "Install Dispatch Watch" shortcut) — copy the app to a local per-user folder,
		// create a Desktop shortcut, and relaunch from there (running off V: directly is
		// what made the app "beaucoup trop lent" in the first place); (2) already running
		// from the local folder — check whether V: has a newer build and, if so, update and
		// relaunch. Either branch that finds work to do calls Environment.Exit and never
		// returns to the caller.
		// Writes to deploy_log.txt next to the exe — EnsureLocalInstall's own Environment.Exit
		// calls bypass Program.cs's crash logger entirely (Exit terminates immediately, no
		// exception involved), so a silent-exit bug in here would otherwise leave zero trace.
		private static void LogDeploy(string message)
		{
			try
			{
				File.AppendAllText(Path.Combine(Application.StartupPath, "deploy_log.txt"),
					DateTime.Now.ToString("s") + " " + message + "\r\n");
			}
			catch { }
		}

		public void EnsureLocalInstall()
		{
			try
			{
				// LocalApplicationData (%LOCALAPPDATA%), not MyDocuments: on this fleet, Documents
				// is OneDrive-redirected (Known Folder Move) — OneDrive's sync client is known to
				// touch local files' LastWriteTime independently of real writes, which silently
				// broke StartApp()'s "is the local copy newer-or-equal than V:?" check (it kept
				// seeing a pre-OCC-migration local ICAO_storedNotams.mdb as "up to date" and never
				// re-downloaded the real one, crashing on the missing Stations_ICAO_IATA table).
				// %LOCALAPPDATA% is explicitly meant for non-roaming, machine-local app data and
				// is never OneDrive-synced, which sidesteps the whole class of problem for a
				// frequently-rewritten binary Access file.
				string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Dispatch Watch");
				string localExe = Path.Combine(localFolder, AppExeName);
				string vExe = Path.Combine(VAppFolder, AppExeName);

				bool alreadyLocal = string.Equals(
					Path.GetFullPath(Application.StartupPath).TrimEnd('\\'),
					Path.GetFullPath(localFolder).TrimEnd('\\'),
					StringComparison.OrdinalIgnoreCase);

				LogDeploy("EnsureLocalInstall: StartupPath=" + Application.StartupPath + " localFolder=" + localFolder +
					" alreadyLocal=" + alreadyLocal + " vExeExists=" + File.Exists(vExe));

				if (!alreadyLocal)
				{
					if (!File.Exists(vExe)) { LogDeploy("V: exe not reachable — running as-is."); return; }
					LogDeploy("First run for this folder — installing to " + localFolder);
					CopyAppFiles(Application.StartupPath, localFolder);
					CreateDesktopShortcut(localExe, localFolder);
					LogDeploy("Install copy done — relaunching from " + localExe);
					Process.Start(new ProcessStartInfo { FileName = localExe, WorkingDirectory = localFolder });
					Environment.Exit(0);
				}
				else if (File.Exists(vExe) && File.GetLastWriteTime(vExe) > File.GetLastWriteTime(localExe))
				{
					LogDeploy("V: exe newer (V:=" + File.GetLastWriteTime(vExe).ToString("s") +
						" local=" + File.GetLastWriteTime(localExe).ToString("s") + ") — updating.");
					CopyAppFiles(VAppFolder, localFolder);
					LogDeploy("Update copy done — relaunching from " + localExe);
					Process.Start(new ProcessStartInfo { FileName = localExe, WorkingDirectory = localFolder });
					Environment.Exit(0);
				}
				else
				{
					LogDeploy("Already local and up to date — continuing normal startup.");
				}
			}
			catch (Exception ex)
			{
				LogDeploy("EXCEPTION: " + ex);
				// deployment/update check is best-effort — must never block the app from starting
			}
		}

		private static void CopyAppFiles(string sourceDir, string destDir)
		{
			Directory.CreateDirectory(destDir);
			foreach (string file in Directory.GetFiles(sourceDir))
			{
				string name = Path.GetFileName(file);
				string ext = Path.GetExtension(file).ToLowerInvariant();
				bool excluded = Array.IndexOf(_deployExcludeExtensions, ext) >= 0;
				if (!excluded)
					foreach (string f in _deployExcludeFiles)
						if (string.Equals(f, name, StringComparison.OrdinalIgnoreCase)) { excluded = true; break; }
				if (excluded) continue;
				File.Copy(file, Path.Combine(destDir, name), true);
			}
		}

		// One-time migration: OCC.mdb used to hold Stations_ICAO_IATA/Runways/Notams_ICAO_CSV
		// in a second Access file with no sync of its own (unlike ICAO_storedNotams.mdb's
		// StartApp/EndApp), which is what left a fresh local install missing it entirely.
		// Folding those tables into ICAO_storedNotams.mdb removes the second file altogether —
		// everything now travels with the one DB that's already synced for every dispatcher.
		// Runs against the shared V: files directly (not Application.StartupPath) so it takes
		// effect exactly once for everyone regardless of which dispatcher happens to launch
		// first or from where — idempotent per table (each of the three is checked and
		// imported independently, not gated behind a single existence check), so it's safe to
		// leave this call in permanently rather than removing it after the first successful
		// run, AND self-healing if an earlier attempt partially failed (e.g. network hiccup
		// mid-import left Runways created but Stations_ICAO_IATA missing — a single "does
		// Runways exist" gate would have skipped the whole migration forever after that,
		// permanently missing Stations_ICAO_IATA). Uses Jet's cross-database "IN 'path'" SQL
		// to import each table's structure and data in one statement rather than
		// hand-rebuilding schemas.
		public void EnsureOccMigrated()
		{
			string vDb = Path.Combine(VAppFolder, "ICAO_storedNotams.mdb");
			string occDb = Path.Combine(VAppFolder, "OCC.mdb");
			if (!File.Exists(vDb) || !File.Exists(occDb)) return;   // V: unreachable, or already merged (OCC.mdb removed once confirmed migrated)

			OleDbConnection conn = null;
			try
			{
				conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= " + vDb);
				conn.Open();

				string[] tables = { "Stations_ICAO_IATA", "Runways", "Notams_ICAO_CSV" };
				foreach (string table in tables)
				{
					bool exists;
					try
					{
						new OleDbCommand("SELECT TOP 1 * FROM [" + table + "]", conn).ExecuteReader().Close();
						exists = true;
					}
					catch { exists = false; }

					if (exists)
					{
						LogDeploy("EnsureOccMigrated: " + table + " already present — skipping.");
						continue;
					}

					try
					{
						new OleDbCommand("SELECT * INTO [" + table + "] FROM [" + table + "] IN '" + occDb + "'", conn).ExecuteNonQuery();
						LogDeploy("EnsureOccMigrated: imported table " + table + ".");
					}
					catch (Exception ex)
					{
						LogDeploy("EnsureOccMigrated: FAILED importing " + table + ": " + ex.Message);
					}
				}
				LogDeploy("EnsureOccMigrated: done.");
			}
			catch (Exception ex)
			{
				LogDeploy("EnsureOccMigrated: EXCEPTION: " + ex);
			}
			finally
			{
				if (conn != null) conn.Close();
			}
		}

		// Late-bound WScript.Shell COM (no project/COM reference needed) — same idiom as the
		// Outlook automation elsewhere in this app (MainForm.Email.cs).
		private static void CreateDesktopShortcut(string targetExePath, string workingDir)
		{
			try
			{
				string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
				string lnkPath = Path.Combine(desktop, "Dispatch Watch.lnk");
				Type shellType = Type.GetTypeFromProgID("WScript.Shell");
				if (shellType == null) return;
				object shell = Activator.CreateInstance(shellType);
				object shortcut = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { lnkPath });
				Type scType = shortcut.GetType();
				scType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { targetExePath });
				scType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, new object[] { workingDir });
				scType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, new object[] { targetExePath });
				scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
			}
			catch { /* a missing shortcut doesn't block the app from running */ }
		}

		// ── User identity (C) ───────────────────────────────────────────────
		private string _currentUserEmail = "";

		// Late-bound Outlook COM, same pattern as the mail-sending code (MainForm.Email.cs).
		// GetExchangeUser().PrimarySmtpAddress is used ahead of AddressEntry.Address because
		// on an Exchange account the plain .Address is often an X.500 legacyExchangeDN string,
		// not a real email address — PrimarySmtpAddress is the reliable one.
		private static string GetCurrentUserEmail()
		{
			try
			{
				Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
				if (outlookType == null) return "";
				object outlook = Activator.CreateInstance(outlookType);
				object ns = outlookType.InvokeMember("GetNamespace", BindingFlags.InvokeMethod, null, outlook, new object[] { "MAPI" });
				object currentUser = ns.GetType().InvokeMember("CurrentUser", BindingFlags.GetProperty, null, ns, null);
				object addrEntry = currentUser.GetType().InvokeMember("AddressEntry", BindingFlags.GetProperty, null, currentUser, null);

				string email = "";
				try
				{
					object exchUser = addrEntry.GetType().InvokeMember("GetExchangeUser", BindingFlags.InvokeMethod, null, addrEntry, null);
					if (exchUser != null)
						email = (string)exchUser.GetType().InvokeMember("PrimarySmtpAddress", BindingFlags.GetProperty, null, exchUser, null);
				}
				catch { /* not an Exchange account — fall through to the plain address below */ }

				if (string.IsNullOrEmpty(email))
					email = (string)addrEntry.GetType().InvokeMember("Address", BindingFlags.GetProperty, null, addrEntry, null);

				return (email ?? "").Trim();
			}
			catch { return ""; }
		}

		public void EnsureUsersTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE Users ([Email] TEXT(255), [FirstSeen] TEXT(50), [LastSeen] TEXT(50))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { /* DB not ready yet */ }
		}

		private static void UpsertCurrentUser(string email)
		{
			if (string.IsNullOrEmpty(email)) return;
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbCommand check = new OleDbCommand("SELECT COUNT(*) FROM Users WHERE Email=?", conn);
				check.Parameters.AddWithValue("?", email);
				int count = (int)check.ExecuteScalar();
				string nowIso = DateTime.Now.ToString("o");
				if (count == 0)
				{
					OleDbCommand ins = new OleDbCommand("INSERT INTO Users ([Email],[FirstSeen],[LastSeen]) VALUES (?,?,?)", conn);
					ins.Parameters.AddWithValue("?", email);
					ins.Parameters.AddWithValue("?", nowIso);
					ins.Parameters.AddWithValue("?", nowIso);
					ins.ExecuteNonQuery();
				}
				else
				{
					OleDbCommand upd = new OleDbCommand("UPDATE Users SET LastSeen=? WHERE Email=?", conn);
					upd.Parameters.AddWithValue("?", nowIso);
					upd.Parameters.AddWithValue("?", email);
					upd.ExecuteNonQuery();
				}
				conn.Close();
			}
			catch { /* identity tracking is best-effort, never blocks startup */ }
		}

		// ── Writer/Reader lock (D) ──────────────────────────────────────────
		// A flat "key|value"-per-line-free text file on V:, next to the .mdb files — same
		// style as MmScheduleCursor.txt (MainForm.MmSchedule.cs). Token is a fresh GUID per
		// claim so a demoted Writer's heartbeat tick can tell "the lock changed out from
		// under me" apart from its own routine heartbeat rewrites.
		private static string LockFilePath { get { return Path.Combine(VAppFolder, "dispatch_watch.lock"); } }
		private const int HeartbeatSeconds = 30;
		private const int StaleLockSeconds = HeartbeatSeconds * 3;   // tolerate a couple of missed ticks before treating a lock as abandoned

		private struct LockInfo
		{
			public bool Found;
			public string Email, Token, Machine;
			public DateTime AcquiredAtUtc, LastHeartbeatUtc;
		}

		public bool IsWriter { get { return _isWriter; } }
		private bool _isWriter;
		private string _writerToken = "";
		private DateTime _lockAcquiredAtUtc;
		private Timer _lockHeartbeatTimer;

		private static LockInfo ReadLock()
		{
			LockInfo info = new LockInfo();
			try
			{
				if (!File.Exists(LockFilePath)) return info;
				string[] parts = File.ReadAllText(LockFilePath).Trim().Split('|');
				if (parts.Length != 5) return info;
				info.Email = parts[0];
				info.Token = parts[1];
				info.Machine = parts[2];
				DateTime acquired, hb;
				DateTime.TryParse(parts[3], CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out acquired);
				DateTime.TryParse(parts[4], CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out hb);
				info.AcquiredAtUtc = acquired;
				info.LastHeartbeatUtc = hb;
				info.Found = true;
			}
			catch { }
			return info;
		}

		private static void WriteLock(string email, string token, DateTime acquiredAtUtc)
		{
			try
			{
				File.WriteAllText(LockFilePath, email + "|" + token + "|" + Environment.MachineName + "|" +
					acquiredAtUtc.ToString("o") + "|" + DateTime.UtcNow.ToString("o"));
			}
			catch { /* couldn't refresh the heartbeat this tick — next tick tries again */ }
		}

		// Runs after identity resolution, before StartApp() downloads the DB. Claims Writer
		// outright when the lock is missing/stale/already ours; otherwise prompts Reader vs
		// Force Writer. Either way, starts the recurring heartbeat that both keeps a Writer's
		// claim alive and detects a forced takeover.
		public void AcquireOrPromptForLock()
		{
			LockInfo existing = ReadLock();
			bool staleOrMissing = !existing.Found || (DateTime.UtcNow - existing.LastHeartbeatUtc).TotalSeconds > StaleLockSeconds;
			bool isOurs = existing.Found && string.Equals(existing.Email, _currentUserEmail, StringComparison.OrdinalIgnoreCase);

			if (staleOrMissing || isOurs)
			{
				ClaimWriter();
			}
			else
			{
				bool forceWriter = ShowLockPromptDialog(existing.Email, existing.AcquiredAtUtc);
				if (forceWriter) ClaimWriter();
				else _isWriter = false;
			}

			_lockHeartbeatTimer = new Timer { Interval = HeartbeatSeconds * 1000 };
			_lockHeartbeatTimer.Tick += LockHeartbeatTick;
			_lockHeartbeatTimer.Start();
		}

		private void ClaimWriter()
		{
			_writerToken = Guid.NewGuid().ToString("N");
			_lockAcquiredAtUtc = DateTime.UtcNow;
			WriteLock(_currentUserEmail, _writerToken, _lockAcquiredAtUtc);
			_isWriter = true;
		}

		private void LockHeartbeatTick(object sender, EventArgs e)
		{
			if (!_isWriter) return;   // Readers just idle — relaunching is how they retry becoming Writer

			LockInfo current = ReadLock();
			if (current.Found && current.Token != _writerToken)
			{
				// Someone forced us out. Emergency-save whatever we have before losing write
				// access, per the confirmed policy — better a slightly-behind V: copy than a
				// silently-lost session's worth of edits.
				UploadDbToV();
				_isWriter = false;
				ApplyReaderModeUi();
				ShowTransientNotice("You've been switched to Reader mode by " + current.Email + ".");
			}
			else
			{
				WriteLock(_currentUserEmail, _writerToken, _lockAcquiredAtUtc);
			}
		}

		public void ReleaseLock()
		{
			if (!_isWriter) return;
			try
			{
				LockInfo l = ReadLock();
				if (l.Found && l.Token == _writerToken) File.Delete(LockFilePath);
			}
			catch { /* next Writer's stale-lock check (90s) recovers from this either way */ }
		}

		private static bool ShowLockPromptDialog(string writerEmail, DateTime acquiredAtUtc)
		{
			using (Form dlg = new Form
			{
				Text = "Dispatch Watch", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false, Width = 460, Height = 190, BackColor = Color.FromArgb(38, 50, 56)
			})
			{
				Label msg = new Label
				{
					Text = "Dispatch Watch is already in use by:\n\n" + writerEmail + "\nsince " + acquiredAtUtc.ToLocalTime().ToString("dd/MM HH:mm") +
						"\n\nContinue as read-only Reader, or take over as Writer?",
					ForeColor = Color.White, Font = new Font("Segoe UI", 9.5f), Dock = DockStyle.Top, Height = 110, TextAlign = ContentAlignment.MiddleCenter
				};
				bool forceWriter = false;
				Button btnReader = new Button { Text = "Continue as Reader", Left = 30, Top = 122, Width = 190, Height = 32 };
				Button btnForce = new Button { Text = "Force Writer", Left = 240, Top = 122, Width = 190, Height = 32, BackColor = Color.IndianRed, ForeColor = Color.White };
				btnReader.Click += delegate { dlg.Close(); };
				btnForce.Click += delegate { forceWriter = true; dlg.Close(); };
				dlg.Controls.Add(btnReader);
				dlg.Controls.Add(btnForce);
				dlg.Controls.Add(msg);
				dlg.ShowDialog();
				return forceWriter;
			}
		}

		private static void ShowTransientNotice(string text)
		{
			Form popup = new Form
			{
				FormBorderStyle = FormBorderStyle.None, StartPosition = FormStartPosition.Manual,
				Width = 380, Height = 70, BackColor = Color.FromArgb(38, 50, 56), TopMost = true, ShowInTaskbar = false
			};
			Rectangle wa = Screen.PrimaryScreen.WorkingArea;
			popup.Location = new Point(wa.Right - popup.Width - 16, wa.Bottom - popup.Height - 16);
			popup.Controls.Add(new Label { Text = text, ForeColor = Color.White, Dock = DockStyle.Fill,
				TextAlign = ContentAlignment.MiddleCenter, Padding = new Padding(10) });
			popup.Show();
			Timer t = new Timer { Interval = 4000 };
			t.Tick += delegate { t.Stop(); popup.Close(); popup.Dispose(); };
			t.Start();
		}

		// Silent counterpart to EndApp()'s upload step (MainForm.cs) — used by the forced-
		// takeover and inactivity-auto-save paths, neither of which should interrupt the
		// dispatcher with a MessageBox the way an explicit app close still does.
		public void UploadDbToV()
		{
			string localPath = Path.Combine(Application.StartupPath, "ICAO_storedNotams.mdb");
			string vDrivePath = Path.Combine(VAppFolder, "ICAO_storedNotams.mdb");
			try
			{
				if (!File.Exists(localPath)) return;
				if (!File.Exists(vDrivePath)) { File.Copy(localPath, vDrivePath, true); return; }
				if (File.GetLastWriteTime(localPath) > File.GetLastWriteTime(vDrivePath))
					File.Copy(localPath, vDrivePath, true);
			}
			catch { /* best-effort */ }
		}

		// Central guard for every DB-write-triggering entry point — see CLAUDE.md for the full
		// list of call sites. Returns false (and tells the dispatcher why) instead of letting a
		// Reader perform an action that would silently be lost (Reader mode never re-uploads).
		public bool EnsureWriterOrWarn()
		{
			if (_isWriter) return true;
			LockInfo l = ReadLock();
			string who = l.Found ? l.Email : "another user";
			MessageBox.Show("Read-only mode — the current writer is " + who + ".\nThis action is disabled while in Reader mode.",
				"Dispatch Watch — Reader mode", MessageBoxButtons.OK, MessageBoxIcon.Information);
			return false;
		}

		// Disables the write-gated controls so Reader mode reads as read-only at a glance,
		// not just as a wall of blocked-action popups from EnsureWriterOrWarn(). Called once
		// right after startup if we opened as a Reader, and again if a live Writer gets
		// forced down mid-session (LockHeartbeatTick).
		public void ApplyReaderModeUi()
		{
			try
			{
				Btn_filterNew.Enabled = false;
				Btn_dbUpdateQuick.Enabled = false;
				Btn_updateDB.Enabled = false;
				Btn_sendReports.Enabled = false;
				Btn_addRecipient.Enabled = false;
				Btn_removeRecipient.Enabled = false;
			}
			catch { /* a control not existing yet (very early startup) shouldn't block the rest */ }
		}

		// ── Inactivity auto-save (E) ────────────────────────────────────────
		private DateTime _lastActivityUtc = DateTime.UtcNow;
		private Timer _idleCheckTimer;

		private class IdleActivityFilter : IMessageFilter
		{
			private const int WM_MOUSEMOVE = 0x0200, WM_LBUTTONDOWN = 0x0201, WM_KEYDOWN = 0x0100;
			private readonly MainForm _form;
			public IdleActivityFilter(MainForm form) { _form = form; }
			public bool PreFilterMessage(ref Message m)
			{
				if (m.Msg == WM_MOUSEMOVE || m.Msg == WM_LBUTTONDOWN || m.Msg == WM_KEYDOWN)
					_form._lastActivityUtc = DateTime.UtcNow;
				return false;
			}
		}

		public void StartIdleAutoSave()
		{
			Application.AddMessageFilter(new IdleActivityFilter(this));
			_idleCheckTimer = new Timer { Interval = 10000 };
			_idleCheckTimer.Tick += delegate
			{
				if (_isWriter && (DateTime.UtcNow - _lastActivityUtc).TotalMinutes >= 2)
				{
					UploadDbToV();
					WriteLock(_currentUserEmail, _writerToken, _lockAcquiredAtUtc);   // heartbeat too, since this proves liveness
					ShowTransientNotice("Auto-saved to V: at " + DateTime.Now.ToString("HH:mm") + " (2 min inactivity).");
					_lastActivityUtc = DateTime.UtcNow;   // don't re-fire every 10s for the rest of the idle period
				}
			};
			_idleCheckTimer.Start();
		}
	}
}
