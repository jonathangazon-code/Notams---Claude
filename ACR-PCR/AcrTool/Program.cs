using System;
using System.IO;
using System.Windows.Forms;

namespace AcrTool
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			// The form is constructed as its own statement, outside Application.Run,
			// so a constructor-time failure still produces a message and a log
			// rather than a window that never appears. Same reasoning as
			// Dispatch Watch's Program.cs.
			MainForm form;
			try
			{
				form = new MainForm();
			}
			catch (Exception ex)
			{
				ReportFatal("starting up", ex);
				return;
			}

			Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e)
			{
				ReportFatal("running", e.Exception);
			};

			AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
			{
				ReportFatal("running", e.ExceptionObject as Exception);
			};

			Application.Run(form);
		}

		private static void ReportFatal(string phase, Exception ex)
		{
			string text = ex == null ? "Unknown error." : ex.ToString();
			string path = Path.Combine(Application.StartupPath, "acr_crash_log.txt");

			try
			{
				File.AppendAllText(path,
					DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  (" + phase + ")" +
					Environment.NewLine + text + Environment.NewLine + Environment.NewLine);
			}
			catch { /* logging must never mask the original error */ }

			MessageBox.Show(
				"ACR/PCR hit an error while " + phase + "." + Environment.NewLine + Environment.NewLine +
				(ex == null ? "" : ex.Message) + Environment.NewLine + Environment.NewLine +
				"Details written to:" + Environment.NewLine + path,
				"ACR / PCR", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
	}
}
