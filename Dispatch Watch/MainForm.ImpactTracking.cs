using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Snapshot of what the engine suggested for a NOTAM at render time (Filter tab
		// pending model only — the dispatcher's final answer isn't known until SUBMIT,
		// so the suggestion has to be captured up front to compare against later).
		private struct NotamSnapshot
		{
			public string Icao, Key, NotamText;
			public List<RwyInfo> Runways;
			public string SuggestedImpact;
			public bool SuggestedSup;
		}
		private Dictionary<int, NotamSnapshot> _pendSnapshot = new Dictionary<int, NotamSnapshot>();

		// Shared on V: (same folder as ICAO_storedNotams.mdb), not Application.StartupPath —
		// since the multi-user deployment gives every dispatcher their own local install
		// (MainForm.Deployment.cs), a local-only log would only ever capture that one
		// dispatcher's corrections, with no way to review the fleet's combined signal.
		// File.AppendAllText opens/closes per write, so concurrent appends from different
		// dispatchers are short enough to not meaningfully collide; a failure here (e.g. V:
		// briefly unreachable) is swallowed by LogImpactCorrection's own try/catch — logging
		// must never block the dispatcher's actual Keep/Impact workflow.
		private static string ImpactCorrectionsLogPath { get { return @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\ImpactCorrections.csv"; } }

		// Formats the exact runway config the engine used for its suggestion (Desig +
		// CatMax), e.g. "05:CAT1; 23:CAT3" — the same RwyInfo list passed into
		// SuggestImpacts, so this is what actually drove the (possibly wrong) decision.
		private static string FormatRunwayConfig(List<RwyInfo> runways)
		{
			if (runways == null || runways.Count == 0) return "";
			List<string> parts = new List<string>();
			foreach (RwyInfo r in runways)
				parts.Add(r.Desig + ":" + (r.CatMax > 0 ? "CAT" + r.CatMax : "none"));
			return string.Join("; ", parts.ToArray());
		}

		private static string CsvField(string s)
		{
			return "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
		}

		// Logs a dispatcher correction of the engine's suggestion — only when the final
		// state actually differs from what was suggested, so the file stays a focused
		// signal of real classification errors rather than every accepted suggestion.
		// finalImpact == "(ignored)" marks a NOTAM rejected outright via Ignore rather than
		// having its impact chip changed.
		private static void LogImpactCorrection(string icao, string key, string notamText,
			List<RwyInfo> runways,
			string suggestedImpact, string finalImpact, bool suggestedSup, bool finalSup)
		{
			if (suggestedImpact == finalImpact && suggestedSup == finalSup) return;   // accepted as-is — nothing to learn from

			try
			{
				List<string> actionParts = new List<string>();
				if (suggestedImpact != finalImpact)
				{
					if (suggestedImpact == "") actionParts.Add("Added " + finalImpact + " (no suggestion)");
					else if (finalImpact == "(ignored)") actionParts.Add("Rejected (Ignored) auto-Keep suggested " + suggestedImpact);
					else if (finalImpact == "") actionParts.Add("Removed suggested " + suggestedImpact);
					else actionParts.Add("Changed " + suggestedImpact + " to " + finalImpact);
				}
				if (suggestedSup != finalSup)
					actionParts.Add(finalSup ? "Added SUP (no suggestion)" : "Removed suggested SUP");
				string action = string.Join("; ", actionParts.ToArray());

				bool isNew = !File.Exists(ImpactCorrectionsLogPath);
				StringBuilder sb = new StringBuilder();
				if (isNew)
					sb.Append("Timestamp,ICAO,NotamKey,RunwayConfig,SuggestedImpact,FinalImpact,SuggestedSup,FinalSup,Action,NotamText\r\n");

				sb.Append(CsvField(DateTime.Now.ToString("s"))).Append(",");
				sb.Append(CsvField(icao)).Append(",");
				sb.Append(CsvField(key)).Append(",");
				sb.Append(CsvField(FormatRunwayConfig(runways))).Append(",");
				sb.Append(CsvField(suggestedImpact)).Append(",");
				sb.Append(CsvField(finalImpact)).Append(",");
				sb.Append(CsvField(suggestedSup ? "Yes" : "")).Append(",");
				sb.Append(CsvField(finalSup ? "Yes" : "")).Append(",");
				sb.Append(CsvField(action)).Append(",");
				sb.Append(CsvField(notamText)).Append("\r\n");

				File.AppendAllText(ImpactCorrectionsLogPath, sb.ToString());
			}
			catch { /* logging must never block the dispatcher's actual workflow */ }
		}
	}
}
