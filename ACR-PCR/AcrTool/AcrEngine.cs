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
	public class AcrEngine
	{
		// aircraft.xml is read in US units throughout: pounds, psi, inches.
		const bool Metric = false;

		Dictionary<string, AircraftEntry> _lib;

		public string LibraryPath { get; private set; }
		public string LibraryVersion { get; private set; }

		public void Load(string aircraftXmlPath)
		{
			_lib = AircraftLibrary.Load(aircraftXmlPath);
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

		/// <summary>ACR at a given gross weight, for all four subgrade categories.</summary>
		public AcrResult Acr(AircraftSpec spec, float grossWeightLb, PavementKind pavement)
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

			ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
			ACRClassLib.clsACR.ACRdata data;

			if (!spec.HasBodyGear)
			{
				data = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1,
				                           main.TyrePressurePsi, x1, y1, Metric);
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

				data = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1, main.TyrePressurePsi, x1, y1,
				                           percent2, wheels2, body.TyrePressurePsi, x2, y2, Metric);
			}

			return Unpack(data);
		}

		AcrResult RigidOneTruck(ACRClassLib.clsACR.PavementType pt, string libraryName, float grossWeightLb)
		{
			AircraftEntry ac = Entry(libraryName);

			int wheels = ac.WheelCount / 2;
			float[] x, y;
			Coords(ac, wheels, out x, out y);

			ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
			return Unpack(runner.CalculateACR(pt, grossWeightLb, ac.MainGearPercent, wheels,
			                                  ac.TyrePressurePsi, x, y, Metric));
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
		public float MaxAllowableWeightLb(AircraftSpec spec, PcrCode pcr, out bool limitedByPavement)
		{
			limitedByPavement = false;

			float mtow = MaxWeightLb(spec);
			if (Acr(spec, mtow, pcr.Pavement).For(pcr.Subgrade) <= pcr.Value)
				return mtow;

			limitedByPavement = true;

			float low = mtow * 0.20f;              // well below any realistic empty weight
			if (Acr(spec, low, pcr.Pavement).For(pcr.Subgrade) > pcr.Value)
				return 0f;

			float high = mtow;
			for (int i = 0; i < 20 && (high - low) > 50f; i++)
			{
				float mid = (low + high) / 2f;
				if (Acr(spec, mid, pcr.Pavement).For(pcr.Subgrade) <= pcr.Value)
					low = mid;
				else
					high = mid;
			}
			return low;
		}

		public const float LbPerKg = 2.2046226f;

		public static float LbToKg(float lb) { return lb / LbPerKg; }
		public static float KgToLb(float kg) { return kg * LbPerKg; }
	}
}
