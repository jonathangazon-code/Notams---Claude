using System;
using System.IO;
using System.Windows.Forms;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		void ExportToPdf(WebBrowser browser, string baseName)
		{
			ExportToPdf(browser.DocumentText, baseName);
		}

		// silent=true (RefreshFlightSchedule's automatic PDF export) skips both the success and
		// failure MessageBox — the caller folds a mention into its own progress-dialog status
		// line instead, so the dispatcher isn't shown two separate "it worked" confirmations.
		// Existing manual "Export PDF" buttons keep silent=false (the default) unchanged.
		bool ExportToPdf(string html, string baseName, bool silent = false)
		{
			string wkhtmlExe = Path.Combine(Application.StartupPath, "wkhtmltopdf.exe");
			if (!File.Exists(wkhtmlExe))
			{
				if (!silent)
					MessageBox.Show(
						"wkhtmltopdf.exe not found in app folder:\n" + wkhtmlExe +
						"\n\nDownload it free from https://wkhtmltopdf.org and place it next to the app.",
						"Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}

			string tempHtml = Path.Combine(Application.StartupPath, "_export_temp.html");
			string pdfName  = DateTime.Now.ToString("yyMMdd") + "-" + baseName + ".pdf";
			string pdfPath  = Path.Combine(
				@"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports",
				pdfName);

			try
			{
				File.WriteAllText(tempHtml, html, System.Text.Encoding.UTF8);

				var proc = new System.Diagnostics.Process();
				proc.StartInfo.FileName        = wkhtmlExe;
				// A very high --dpi with no matching --zoom shrinks all CSS px-sized content
				// (text, table borders, the AIP SUP checkboxes) tiny and blurry on the page —
				// text in a PDF is vector-drawn regardless of DPI, so a plain 96 DPI (the
				// default wkhtmltopdf/browser assumption) renders crisp at any zoom level.
				// --enable-local-file-access: newer wkhtmltopdf builds refuse to load local
				// file:// subresources (e.g. the runway-diagram PNGs in %TEMP%, MainForm.Conflict.cs's
				// BuildRwyDiagramImageTag) that live outside the source HTML's own directory
				// (_export_temp.html is next to the exe) — without this flag they render as
				// blank/broken images.
				proc.StartInfo.Arguments       = "--enable-local-file-access --dpi 96 --disable-smart-shrinking --image-quality 100 \"" + tempHtml + "\" \"" + pdfPath + "\"";
				proc.StartInfo.UseShellExecute = false;
				proc.StartInfo.CreateNoWindow  = true;
				// The runway-diagram PNGs still aren't showing in exported PDFs despite the
				// file:// URI fix and --enable-local-file-access (the sibling airportDocu
				// project hit the exact same wall with its ASL-logo image), and guessing a
				// third fix blind isn't productive — wkhtmltopdf normally prints exactly why a
				// resource failed to load on stderr, but that was being silently discarded
				// (CreateNoWindow with no redirection). Logging it next to the exe surfaces the
				// real reason.
				proc.StartInfo.RedirectStandardError  = true;
				proc.StartInfo.RedirectStandardOutput = true;
				proc.Start();
				string stderrOutput = proc.StandardError.ReadToEnd();
				string stdoutOutput = proc.StandardOutput.ReadToEnd();
				proc.WaitForExit();

				string logPath = Path.Combine(Application.StartupPath, "wkhtmltopdf_log.txt");
				File.WriteAllText(logPath, "=== stdout ===\r\n" + stdoutOutput + "\r\n=== stderr ===\r\n" + stderrOutput);

				File.Delete(tempHtml);

				if (File.Exists(pdfPath))
				{
					if (!silent) MessageBox.Show("PDF exported:\n" + pdfPath, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return true;
				}
				else
				{
					if (!silent) MessageBox.Show("Export failed — wkhtmltopdf returned no file.", "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
			catch (Exception ex)
			{
				if (!silent) MessageBox.Show("Export error:\n" + ex.Message, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}
	}
}
