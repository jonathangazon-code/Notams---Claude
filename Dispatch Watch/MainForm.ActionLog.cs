using System;
using System.IO;
using System.Text;
using System.Threading;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Shared, append-only audit trail of every write action any dispatcher performs —
		// one line per action, "yyyy-MM-dd HH:mm:ss | user@email | Action text". Lives next to
		// ICAO_storedNotams.mdb on V: (VAppFolder, MainForm.Deployment.cs), NOT
		// Application.StartupPath, for the same reason every other shared record in this app
		// does (Users table, dispatch_watch.lock, MmScheduleCursor.txt): every dispatcher runs
		// their own local install, so a per-instance log would only ever capture that one
		// dispatcher's actions instead of the whole team's.
		private static readonly string ActionLogPath = Path.Combine(VAppFolder, "UserActions_log.txt");

		// Must never add perceptible latency to a button click — this is an audit trail, not
		// data the app depends on. LogUserAction only ever queues the line and returns
		// immediately; the actual network write to V: happens on a background ThreadPool
		// thread. A failed/slow/unreachable V: write is swallowed (try/catch) and simply
		// drops that line rather than blocking or popping an error — same "best effort,
		// never in the way" spirit as the inactivity auto-save's background upload.
		// _actionLogLock only serializes the actual file write (multiple background writes
		// racing each other within this one process); it says nothing about other
		// dispatchers' processes writing to the same UNC path at the same moment — File.
		// AppendAllText's own open/write/close sequence keeps concurrent cross-process
		// appends from corrupting each other's lines in practice, and an occasional lost
		// line under real contention is an acceptable trade-off for a log that must not
		// slow the app down or introduce a shared write-lock of its own.
		private static readonly object _actionLogLock = new object();

		public void LogUserAction(string action)
		{
			string user = string.IsNullOrEmpty(_currentUserEmail) ? "(unknown)" : _currentUserEmail;
			string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " | " + user + " | " + action;

			ThreadPool.QueueUserWorkItem(delegate
			{
				try
				{
					lock (_actionLogLock)
					{
						File.AppendAllText(ActionLogPath, line + "\r\n", Encoding.UTF8);
					}
				}
				catch { }
			});
		}
	}
}
