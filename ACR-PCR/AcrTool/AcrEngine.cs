using System;
using System.Collections.Generic;

namespace AcrTool
{
	/// <summary>Result of one ACR evaluation, per subgrade category.</summary>
	public class AcrResult
	{
		public float A, B, C, D;

		public float For(char subgrade)
		{
			switch (char.ToUpperInvariant(subgrade))
			{
				case 'A': return A;
				case 'B': return B;
				case 'C': return C;
				case 'D': return D;
				default: throw new ArgumentException("Unknown subgrade category: " + subgrade);
			}
		}
	}

	/// <summary>
	/// Drives ACRClassLib.dll (the FAA ACR engine) from gear data read out of
	/// aircraft.xml by AircraftLibrary.
	///
	/// Which wheels to use, and what share of the weight they carry, is not
	/// something the engine decides - the FAA API document is explicit that the
	/// calling program must determine it. Every rule below is copied from the
	/// FAA's own ICAO-ACR program (ACRClassDriver/Form1_ICAO.vb) rather than
	/// derived independently, so results can be compared against it directly.
	/// </summary>
	public class AcrEngine : IRatingEngine
	{
		// aircraft.xml is read in US units throughout: pounds, psi, inches.
		const bool Metric = false;

		// ACRClassLib keeps its working state in VB modules - gICAOCodeIndex,
		// gPavementType, gStrainTarget and friends - which are static, shared by
		// every instance. It is therefore not thread-safe, and two concurrent
		// CalculateACR calls would corrupt each other's results silently.
		//
		// That rules out evaluating the four aircraft in parallel, and it also
		// matters here: the evaluation runs on a worker thread while the grid's
		// "check a weight" cell can fire another call from the UI thread. This lock
		// serialises every entry into the library.
		static readonly object _libLock = new object();

		/// <summary>
		/// The same lock, for anything else that calls into ACRClassLib directly -
		/// the self-test does, from the UI thread, and could otherwise land in the
		/// middle of a running evaluation.
		/// </summary>
		public static object SyncRoot { get { return _libLock; } }

		// ---- timing diagnostics ------------------------------------------------
		//
		// Two rounds of optimising by reasoning about call counts did not match what
		// the app actually does, so this records what really happens: one line per
		// call into the library, with the time it took. Written next to the exe,
		// rewritten on each run. Cheap enough to leave on.

		static readonly System.Text.StringBuilder _timing = new System.Text.StringBuilder();
		static int _calls;
		static double _totalMs;

		static void Timed(string what, System.Action body)
		{
			System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
			body();
			sw.Stop();

			_calls++;
			_totalMs += sw.Elapsed.TotalMilliseconds;
			lock (_timing)
			{
				if (_timing.Length < 400000)
					_timing.AppendLine(string.Format(
						System.Globalization.CultureInfo.InvariantCulture,
						"{0,8:0.0} ms  {1}", sw.Elapsed.TotalMilliseconds, what));
			}
		}

		/// <summary>Call count and total solver time since the process started.</summary>
		public static string TimingSummary()
		{
			return string.Format(System.Globalization.CultureInfo.InvariantCulture,
				"{0} solver calls, {1:0.0} s total", _calls, _totalMs / 1000.0);
		}

		/// <summary>Writes the per-call log next to the exe. Never throws.</summary>
		public static void DumpTiming(string path)
		{
			try
			{
				lock (_timing)
				{
					System.IO.File.WriteAllText(path,
						"ACR solver timing - " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine +
						TimingSummary() + Environment.NewLine + Environment.NewLine + _timing);
				}
			}
			catch { /* diagnostics must never break the app */ }
		}

		Dictionary<string, AircraftEntry> _lib;

		// One layered-elastic solve is expensive, and the same weight gets asked
		// for repeatedly - the row shows ACR at max weight, then the search for the
		// allowable weight starts by evaluating that same point again, and every
		// re-render (unit toggle, overload toggle) repeats the lot.
		readonly Dictionary<string, AcrResult> _cache = new Dictionary<string, AcrResult>();

		public string LibraryPath { get; private set; }
		public string LibraryVersion { get; private set; }

		public void Load(string aircraftXmlPath)
		{
			_lib = AircraftLibrary.Load(aircraftXmlPath);
			_cache.Clear();
			LibraryPath = aircraftXmlPath;
			LibraryVersion = AircraftLibrary.Version(aircraftXmlPath);
		}

		AircraftEntry Entry(string libraryName)
		{
			if (_lib == null)
				throw new InvalidOperationException("The aircraft library has not been loaded.");

			AircraftEntry e;
			if (!_lib.TryGetValue(libraryName, out e))
				throw new InvalidOperationException(
					"Aircraft \"" + libraryName + "\" is not present in this aircraft.xml.");
			return e;
		}

		/// <summary>Maximum take-off weight from the library, in pounds.</summary>
		public float MaxWeightLb(AircraftSpec spec)
		{
			return Entry(spec.LibraryName).GrossWeightLb;
		}

		/// <summary>Tyre (inflation) pressure from the library, in psi.</summary>
		public float TyrePressurePsi(AircraftSpec spec)
		{
			return Entry(spec.LibraryName).TyrePressurePsi;
		}

		/// <summary>
		/// Takes the first <paramref name="wheels"/> coordinates, keeping the
		/// 1-based layout the engine expects (slot 0 present but unused).
		/// </summary>
		static void Coords(AircraftEntry ac, int wheels, out float[] x, out float[] y)
		{
			x = new float[wheels + 1];
			y = new float[wheels + 1];
			for (int i = 0; i <= wheels && i < ac.X.Length; i++)
			{
				x[i] = ac.X[i];
				y[i] = ac.Y[i];
			}
		}

		/// <summary>
		/// Which wheels define the strain evaluation grid, as the FAA driver builds
		/// it: the mean lateral coordinate is taken, and every wheel at or beyond it
		/// is included - in practice one side of a symmetric gear.
		///
		/// The library does not decide this; the API document is explicit that the
		/// calling program must, and that including every wheel "would take much
		/// longer" for an insignificant difference in the result. Passing nothing,
		/// as this did before, paid exactly that price.
		///
		/// 1-based to match the coordinate arrays: slot 0 is unused.
		/// </summary>
		static int[] StrainGrid(float[] x, int wheels)
		{
			int[] sw = new int[wheels + 1];
			if (wheels <= 0) return sw;

			double sum = 0;
			for (int i = 1; i <= wheels; i++) sum += x[i];
			double mean = sum / wheels;

			for (int i = 1; i <= wheels; i++) sw[i] = x[i] >= mean ? 1 : 0;
			return sw;
		}

		/// <summary>ACR at a given gross weight, for all four subgrade categories.</summary>
		public AcrResult Acr(AircraftSpec spec, float grossWeightLb, PavementKind pavement)
		{
			string key = spec.Display + "|" + (int)pavement + "|" + Math.Round(grossWeightLb);
			AcrResult hit;
			if (_cache.TryGetValue(key, out hit)) return hit;

			AcrResult computed = ComputeAcr(spec, grossWeightLb, pavement);
			_cache[key] = computed;
			return computed;
		}

		AcrResult ComputeAcr(AircraftSpec spec, float grossWeightLb, PavementKind pavement)
		{
			ACRClassLib.clsACR.PavementType pt = pavement == PavementKind.Rigid
				? ACRClassLib.clsACR.PavementType.Rigid
				: ACRClassLib.clsACR.PavementType.Flexible;

			if (pavement == PavementKind.Rigid)
			{
				// Rigid ACR loads one truck only:
				// wheels_number = libNWheels / 2, percent_gw = libMGpcntPCN.
				//
				// On the 747 the wing and body trucks take the same share at the
				// same pressure and differ only in spacing, so instead of guessing
				// which is "most demanding" both are run and the higher value kept
				// for each subgrade.
				AcrResult wing = RigidOneTruck(pt, spec.LibraryName, grossWeightLb);
				if (!spec.HasBodyGear) return wing;

				AcrResult body = RigidOneTruck(pt, spec.BellyLibraryName, grossWeightLb);
				AcrResult worst = new AcrResult();
				worst.A = Math.Max(wing.A, body.A);
				worst.B = Math.Max(wing.B, body.B);
				worst.C = Math.Max(wing.C, body.C);
				worst.D = Math.Max(wing.D, body.D);
				return worst;
			}

			// Flexible ACR accounts for every wheel of the main landing gear.
			AircraftEntry main = Entry(spec.LibraryName);

			int wheels1 = main.WheelCount;
			float percent1 = main.MainGearPercent * 2f;   // entry covers one side; x2 for the pair
			float[] x1, y1;
			Coords(main, wheels1, out x1, out y1);

			int[] sw1 = StrainGrid(x1, wheels1);

			ACRClassLib.clsACR.ACRdata data;

			if (!spec.HasBodyGear)
			{
				ACRClassLib.clsACR.ACRdata d1 = default(ACRClassLib.clsACR.ACRdata);
				Timed(spec.Display + " flex 1-gear " + wheels1 + "w " + Math.Round(grossWeightLb) + "lb",
					delegate
					{
						lock (_libLock)
						{
							ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
							d1 = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1,
							                         main.TyrePressurePsi, x1, y1, sw1, Metric);
						}
					});
				data = d1;
			}
			else
			{
				// Wing gear and body gear go in as two separate gears, the way
				// Form1_ICAO.vb handles the 747 and A380.
				AircraftEntry body = Entry(spec.BellyLibraryName);

				int wheels2 = body.WheelCount;
				float percent2 = body.MainGearPercent * 2f;
				float[] x2, y2;
				Coords(body, wheels2, out x2, out y2);
				int[] sw2 = StrainGrid(x2, wheels2);

				ACRClassLib.clsACR.ACRdata d2 = default(ACRClassLib.clsACR.ACRdata);
				Timed(spec.Display + " flex 2-gear " + wheels1 + "+" + wheels2 + "w " + Math.Round(grossWeightLb) + "lb",
					delegate
					{
						lock (_libLock)
						{
							ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
							d2 = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1, main.TyrePressurePsi, x1, y1,
							                         percent2, wheels2, body.TyrePressurePsi, x2, y2, sw1, sw2, Metric);
						}
					});
				data = d2;
			}

			return Unpack(data);
		}

		AcrResult RigidOneTruck(ACRClassLib.clsACR.PavementType pt, string libraryName, float grossWeightLb)
		{
			AircraftEntry ac = Entry(libraryName);

			int wheels = ac.WheelCount / 2;
			float[] x, y;
			Coords(ac, wheels, out x, out y);

			// No SW here: the rigid path already loads a single truck, and the FAA
			// driver likewise passes the overload without it.
			ACRClassLib.clsACR.ACRdata data = default(ACRClassLib.clsACR.ACRdata);
			Timed(libraryName + " rigid " + wheels + "w " + Math.Round(grossWeightLb) + "lb",
				delegate
				{
					lock (_libLock)
					{
						ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
						data = runner.CalculateACR(pt, grossWeightLb, ac.MainGearPercent, wheels,
						                           ac.TyrePressurePsi, x, y, Metric);
					}
				});
			return Unpack(data);
		}

		/// <summary>
		/// Reads the ACRdata arrays, which are length 5 with slot 0 unused and run
		/// backwards: index 1 is subgrade D and index 4 is subgrade A.
		/// </summary>
		static AcrResult Unpack(ACRClassLib.clsACR.ACRdata data)
		{
			AcrResult r = new AcrResult();
			r.D = data.libACR[1];
			r.C = data.libACR[2];
			r.B = data.libACR[3];
			r.A = data.libACR[4];
			return r;
		}

		/// <summary>
		/// Heaviest weight whose ACR still fits the given PCR, capped at MTOW.
		/// Returns MTOW when the pavement is not the limit, and 0 when the
		/// aircraft does not fit even at the lowest weight considered.
		///
		/// ACR rises monotonically with weight, so a plain bisection works. Each
		/// step runs the layered-elastic solver, which is not cheap, so it stops
		/// at 50 lb - far finer than any operational use of the answer.
		/// </summary>
		public float MaxAllowableWeightLb(AircraftSpec spec, PavementCode pcr, float limit, out bool limitedByPavement)
		{
			limitedByPavement = false;

			float mtow = MaxWeightLb(spec);
			if (Acr(spec, mtow, pcr.Pavement).For(pcr.Subgrade) <= limit)
				return mtow;

			limitedByPavement = true;

			float low = mtow * 0.20f;              // well below any realistic empty weight
			float fLow = Acr(spec, low, pcr.Pavement).For(pcr.Subgrade) - limit;
			if (fLow > 0f)
				return 0f;                         // unusable at any sensible weight

			float high = mtow;
			float fHigh = Acr(spec, high, pcr.Pavement).For(pcr.Subgrade) - limit;

			// ACR is very nearly linear in weight, so interpolating between the
			// bracket ends lands close to the answer immediately. Measured against
			// a fine reference on curves from linear to strongly convex, this needs
			// about 8 solves where plain bisection needed 21, for the same accuracy
			// (~50 lb) - and start-up time was dominated by that count.
			//
			// The classic weakness of false position, one end going stale, does not
			// bite here because the loop stops on how far the estimate moved rather
			// than on how wide the bracket still is. The probe is kept strictly
			// inside the bracket, and the iteration cap bounds the worst case.
			float tolerance = Math.Max(20f, mtow * 0.0002f);
			float estimate = low;
			float previous = float.NaN;

			for (int i = 0; i < 15; i++)
			{
				float span = fHigh - fLow;
				if (Math.Abs(span) < 1e-9f) break;

				estimate = high - fHigh * (high - low) / span;

				float margin = (high - low) * 0.01f;
				if (estimate < low + margin) estimate = low + margin;
				if (estimate > high - margin) estimate = high - margin;

				float fEstimate = Acr(spec, estimate, pcr.Pavement).For(pcr.Subgrade) - limit;

				if (fEstimate <= 0f) { low = estimate; fLow = fEstimate; }
				else                 { high = estimate; fHigh = fEstimate; }

				if (!float.IsNaN(previous) && Math.Abs(estimate - previous) <= tolerance) break;
				previous = estimate;
			}

			// Return the heaviest weight known to fit, never the last probe, which
			// may sit a few pounds above the limit.
			return low;
		}

		// ---- IRatingEngine ----------------------------------------------------

		public string RatingName { get { return "ACR"; } }

		/// <summary>
		/// No minimum weight: aircraft.xml publishes only the maximum weight, and
		/// the ACR engine computes at any weight rather than from a table. Zero
		/// means "not published", and the row shows a dash.
		/// </summary>
		public float MinWeightLb(AircraftSpec spec) { return 0f; }

		public RatingMethod Method { get { return RatingMethod.Acr; } }
		public bool Ready { get { return _lib != null; } }
		public string NotReadyReason { get { return Ready ? null : "The aircraft library has not been loaded."; } }

		public string Provenance
		{
			get
			{
				return "ACR computed by the FAA ICAO-ACR engine (ACRClassLib.dll); gear geometry from the FAA "
					+ "aircraft library (aircraft.xml, version " + (LibraryVersion ?? "not loaded") + ").";
			}
		}

		public float Rating(AircraftSpec spec, float weightLb, PavementCode code)
		{
			return Acr(spec, weightLb, code.Pavement).For(code.Subgrade);
		}

		public const float LbPerKg = 2.2046226f;

		public static float LbToKg(float lb) { return lb / LbPerKg; }
		public static float KgToLb(float kg) { return kg * LbPerKg; }
	}
}
