using System;
using System.IO;
using System.Windows.Forms;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		void ExportToPdf(WebBrowser browser, string baseName)
		{
			string wkhtmlExe = Path.Combine(Application.StartupPath, "wkhtmltopdf.exe");
			if (!File.Exists(wkhtmlExe))
			{
				MessageBox.Show(
					"wkhtmltopdf.exe not found in app folder:\n" + wkhtmlExe +
					"\n\nDownload it free from https://wkhtmltopdf.org and place it next to the app.",
					"Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			string tempHtml = Path.Combine(Application.StartupPath, "_export_temp.html");
			string pdfName  = DateTime.Now.ToString("yyMMdd") + "-" + baseName + ".pdf";
			string pdfPath  = Path.Combine(
				@"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports",
				pdfName);

			try
			{
				File.WriteAllText(tempHtml, browser.DocumentText, System.Text.Encoding.UTF8);

				var proc = new System.Diagnostics.Process();
				proc.StartInfo.FileName        = wkhtmlExe;
				// A very high --dpi with no matching --zoom shrinks all CSS px-sized content
				// (text, table borders, the AIP SUP checkboxes) tiny and blurry on the page —
				// text in a PDF is vector-drawn regardless of DPI, so a plain 96 DPI (the
				// default wkhtmltopdf/browser assumption) renders crisp at any zoom level.
				proc.StartInfo.Arguments       = "--dpi 96 --disable-smart-shrinking --image-quality 100 \"" + tempHtml + "\" \"" + pdfPath + "\"";
				proc.StartInfo.UseShellExecute = false;
				proc.StartInfo.CreateNoWindow  = true;
				proc.Start();
				proc.WaitForExit();

				File.Delete(tempHtml);

				if (File.Exists(pdfPath))
					MessageBox.Show("PDF exported:\n" + pdfPath, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
				else
					MessageBox.Show("Export failed — wkhtmltopdf returned no file.", "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Export error:\n" + ex.Message, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
