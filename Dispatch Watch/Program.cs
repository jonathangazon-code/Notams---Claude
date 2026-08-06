/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 24-02-21
 * Time: 17:16
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Windows.Forms;

namespace ICAO_CSV
{
	/// <summary>
	/// Class with program entry point.
	/// </summary>
	internal sealed class Program
	{
		/// <summary>
		/// Program entry point.
		/// </summary>
		[STAThread]
		private static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			// The app was seen exiting silently (taskbar icon flashes then vanishes) right
			// after StartApp()'s DB download, with no dialog and no log — every new code path
			// from that region is already try/catch-guarded, so the actual cause has to come
			// from somewhere that currently has no visibility at all. new MainForm() runs as
			// the argument to Application.Run BEFORE that call's own message-loop exception
			// handling exists, so an exception thrown during the constructor bypasses
			// Application.ThreadException entirely and just unwinds out of Main() — this is
			// exactly the shape of "flashes then disappears" (the window handle/taskbar icon
			// exists partway through InitializeComponent(), then everything unwinds). Splitting
			// construction out and wrapping it, plus the two handlers below for anything during
			// the later message loop, means the next crash writes a real stack trace to
			// crash_log.txt next to the exe instead of vanishing silently.
			Application.ThreadException += (s, e) => LogCrash(e.Exception);
			AppDomain.CurrentDomain.UnhandledException += (s, e) => LogCrash(e.ExceptionObject as Exception);

			try
			{
				MainForm form = new MainForm();
				Application.Run(form);
			}
			catch (Exception ex)
			{
				LogCrash(ex);
				MessageBox.Show("Dispatch Watch failed to start:\n\n" + ex.Message +
					"\n\nDetails written to crash_log.txt next to the exe.",
					"Dispatch Watch — Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private static void LogCrash(Exception ex)
		{
			try
			{
				string path = Path.Combine(Application.StartupPath, "crash_log.txt");
				string text = "=== " + DateTime.Now.ToString("s") + " ===\r\n" + (ex != null ? ex.ToString() : "(null exception)") + "\r\n\r\n";
				File.AppendAllText(path, text);
			}
			catch { /* if we can't even log, there's nothing more to do here */ }
		}
	}
}
