using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// IATA carrier code -> operator callsign prefix, used only when a Movement Manager
		// message doesn't carry an explicit <callsign> (see LoadMmSupplementalFlights) — a
		// bare 2-letter IATA carrierCode (e.g. "3V") is shared by several different
		// operators, so it can't be trusted on its own. Small/extensible, same idea as
		// _callsignPrefixFilters and the aircraft-type exclusion list in SuggestImpacts.
		private static readonly Dictionary<string, string> _mmCarrierPrefixMap = new Dictionary<string, string>
		{
			{ "3V", "TAY" }, { "QY", "BCS" }, { "5O", "FPO" }, { "5H", "ABR" }, { "FX", "FDX" }
		};

		// Shared on V: (VAppFolder, MainForm.Deployment.cs), not Application.StartupPath-relative:
		// RefreshFlightSchedule only ever runs as Writer (EnsureWriterOrWarn — one instance at a
		// time), so there's no write-contention risk, and sharing it means whichever dispatcher
		// is Writer today resumes from wherever the previous Writer actually left off instead of
		// a cold per-machine cursor forcing a full backlog re-scan.
		private static string MmScheduleCursorPath { get { return Path.Combine(VAppFolder, "MmScheduleCursor.txt"); } }

		private struct MmCursor { public DateTime FolderDate; public DateTime FileTime; }

		private static MmCursor LoadMmCursor()
		{
			try
			{
				if (File.Exists(MmScheduleCursorPath))
				{
					string[] parts = File.ReadAllText(MmScheduleCursorPath).Trim().Split('|');
					if (parts.Length == 2)
					{
						DateTime folderDate = DateTime.ParseExact(parts[0], "yyyyMMdd", CultureInfo.InvariantCulture);
						DateTime fileTime = DateTime.Parse(parts[1], CultureInfo.InvariantCulture,
							DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
						return new MmCursor { FolderDate = folderDate, FileTime = fileTime };
					}
				}
			}
			catch { }
			// First run ever: start two days back — a cheap safety margin catching an
			// update filed just before midnight about a today/future flight, without
			// reprocessing the entire historical backlog.
			return new MmCursor { FolderDate = DateTime.UtcNow.Date.AddDays(-2), FileTime = DateTime.MinValue };
		}

		private static void SaveMmCursor(MmCursor cursor)
		{
			try
			{
				File.WriteAllText(MmScheduleCursorPath,
					cursor.FolderDate.ToString("yyyyMMdd") + "|" + cursor.FileTime.ToString("o"));
			}
			catch { }
		}

		// One flight-instance's latest known state, keyed by carrierCode|flightNumber|originDate.
		private class MmFlightState
		{
			public string Reg = "";
			public string Std = "";
			public string Sta = "";
			public string Origin = "";
			public string Dest = "";
			public string FldT = "";
			public bool Cancelled;
		}

		// Deterministic negative synthetic ID for an MM flight-instance key, kept in a
		// distinct sub-range from the CSV fallback's -csvId (always small-magnitude,
		// derived from FPM's own numeric booking IDs) so the two synthetic ID spaces can
		// never collide.
		private static int MmSyntheticId(string key)
		{
			int hash = 0;
			unchecked { foreach (char c in key) hash = hash * 31 + c; }
			int bucket = Math.Abs(hash) % 900000000;
			return -(1000000000 + bucket);
		}

		// Parses Movement Manager flightInfo_V2 XML messages (dropped continuously into
		// per-day folders under _mmMessagesPath by the exporting system) into supplemental
		// FlightSchedule rows — the baseline source for the 7-day window: superseded by the
		// webservice (wsKeys) once a flight appears there, and itself superseding the CSV
		// fallback. Resumes from the persisted cursor rather than reprocessing the backlog
		// on every refresh.
		private List<object[]> LoadMmSupplementalFlights(HashSet<string> wsKeys, HashSet<int> mmIdsSeen, out HashSet<string> mmKeys)
		{
			mmKeys = new HashSet<string>();
			List<object[]> result = new List<object[]>();
			Dictionary<string, MmFlightState> byKey = new Dictionary<string, MmFlightState>();

			if (string.IsNullOrEmpty(_mmMessagesPath) || !Directory.Exists(_mmMessagesPath)) return result;

			MmCursor cursor = LoadMmCursor();
			MmCursor newCursor = cursor;
			DateTime windowEnd = DateTime.UtcNow.Date.AddDays(7);

			// Probing candidate day-folder names directly (cursor.FolderDate .. today) instead
			// of Directory.GetDirectories(_mmMessagesPath) — the latter lists the share's
			// *entire* archived history over the network just to filter it down to the last
			// few days client-side, which dominated the 4+ minute stalls seen in practice.
			List<KeyValuePair<DateTime, string>> dayFolders = new List<KeyValuePair<DateTime, string>>();
			for (DateTime d = cursor.FolderDate; d <= DateTime.UtcNow.Date; d = d.AddDays(1))
			{
				string candidate = Path.Combine(_mmMessagesPath, d.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
				if (Directory.Exists(candidate)) dayFolders.Add(new KeyValuePair<DateTime, string>(d, candidate));
			}

			foreach (KeyValuePair<DateTime, string> dayFolder in dayFolders)
			{
				DateTime folderDate = dayFolder.Key;
				string[] files;
				try { files = Directory.GetFiles(dayFolder.Value, "*.xml"); } catch { continue; }

				List<KeyValuePair<DateTime, string>> ordered = new List<KeyValuePair<DateTime, string>>();
				foreach (string f in files)
				{
					DateTime wt;
					try { wt = File.GetLastWriteTimeUtc(f); } catch { continue; }
					if (folderDate == cursor.FolderDate && wt <= cursor.FileTime) continue;   // already processed
					ordered.Add(new KeyValuePair<DateTime, string>(wt, f));
				}
				ordered.Sort((a, b) => a.Key.CompareTo(b.Key));
				if (ordered.Count == 0) continue;

				// Each file's XDocument.Load is an independent UNC round-trip — the dominant
				// cost during a catch-up is network latency, not parsing, so reading files
				// concurrently (bounded) cuts wall-clock time roughly by the parallelism
				// factor instead of paying every round-trip one at a time. The parsed <flight>
				// elements are only collected here, not applied to the shared byKey dictionary
				// — that happens in a second, sequential pass below, in write-time order, so
				// "the latest message per key wins" stays deterministic regardless of which
				// parallel read happens to finish first.
				List<XElement>[] parsedPerFile = new List<XElement>[ordered.Count];
				Parallel.For(0, ordered.Count, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i =>
				{
					try
					{
						XDocument doc = XDocument.Load(ordered[i].Value);
						// Most messages carry a single <flight>, but the twice-daily ACCLAIM
						// batch files (~05:30/16:00, ~150-200KB) carry many sibling <flight>
						// elements spanning several days each — the only real source of
						// visibility beyond day+2, since the small per-flight delta messages
						// are mostly near-term changes. Using Element() (first child only)
						// here silently dropped ~99% of a batch file's content; Elements()
						// walks all of them.
						parsedPerFile[i] = doc.Root != null
							? new List<XElement>(doc.Root.Elements(FsNs + "flight"))
							: new List<XElement>();
					}
					catch { parsedPerFile[i] = new List<XElement>(); }   // one bad/partial file shouldn't abort the whole catch-up
				});

				for (int i = 0; i < ordered.Count; i++)
				{
					foreach (XElement flight in parsedPerFile[i])
					{
						string carrier      = El(flight, FsNs + "carrierCode");
						string flightNumber = El(flight, FsNs + "flightNumber");
						string originDate   = El(flight, FsNs + "originDate").TrimEnd('Z');

						if (carrier != "" && flightNumber != "" && originDate != "")
						{
							string key = carrier + "|" + flightNumber + "|" + originDate;
							MmFlightState state;
							if (!byKey.TryGetValue(key, out state)) { state = new MmFlightState(); byKey[key] = state; }

							if (El(flight, FsNs + "action") == "C")
							{
								state.Cancelled = true;
							}
							else
							{
								state.Cancelled = false;
								state.Reg    = El(flight, FsNs + "aircraftRegistration");
								state.Std    = El(flight, FsNs + "scheduledDepartureTime");
								state.Sta    = El(flight, FsNs + "scheduledArrivalTime");
								state.Origin = NormalizeIata(El(flight, FsNs + "departureAerodrome"));
								state.Dest   = NormalizeIata(El(flight, FsNs + "arrivalAerodrome"));
								state.FldT   = originDate;
							}
						}
					}
					newCursor.FolderDate = folderDate;
					newCursor.FileTime = ordered[i].Key;
				}
			}

			SaveMmCursor(newCursor);

			foreach (KeyValuePair<string, MmFlightState> kv in byKey)
			{
				MmFlightState state = kv.Value;
				if (state.Cancelled || state.Std == "") continue;

				string[] keyParts = kv.Key.Split('|');
				string carrier = keyParts[0], flightNumber = keyParts[1];

				// Deliberately NOT using the message's own <callsign> here: it turns out to
				// sometimes hold the tactical/rotation callsign (e.g. "TAY6BS") rather than
				// the flight-number-based one the webservice uses (e.g. "TAY4337" for the
				// exact same leg) — confirmed by cross-checking a live duplicate pair
				// against the source XML. Using it created the same kind of duplicate the
				// CSV's ATC-field callsign once did, so the callsign is always rebuilt from
				// carrierCode + flightNumber via the operator map instead, for consistency.
				string prefix;
				if (!_mmCarrierPrefixMap.TryGetValue(carrier, out prefix)) continue;   // can't attribute — skip
				string callsign = prefix + flightNumber;
				if (!CallsignAllowed(callsign)) continue;

				DateTime stdDt;
				if (!DateTime.TryParse(state.Std, CultureInfo.InvariantCulture,
					DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out stdDt)) continue;
				if (stdDt < DateTime.UtcNow.Date || stdDt > windowEnd) continue;

				string mergeKey = MergeKey(callsign, state.Origin, state.Dest, state.FldT);
				int fltlegId = MmSyntheticId(kv.Key);
				mmIdsSeen.Add(fltlegId);
				mmKeys.Add(mergeKey);
				if (wsKeys.Contains(mergeKey)) continue;   // webservice already covers this flight

				result.Add(new object[] { fltlegId, state.FldT, callsign, state.Reg.Replace("-", "").ToUpper(),
					state.Std, "", "MM", state.Sta, state.Origin, state.Dest });
			}

			return result;
		}
	}
}
