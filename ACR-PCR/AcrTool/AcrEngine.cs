using System;
using System.Collections.Generic;
using System.IO;

namespace AcrTool
{
	/// <summary>
	/// Result of one ACR evaluation.
	/// </summary>
	public class AcrResult
	{
		/// <summary>ACR for subgrade A, B, C, D.</summary>
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
	/// Wraps the FAA libraries:
	///   ACClassLib.clsAC   - reads aircraft.xml into gear geometry (US units)
	///   ACRClassLib.clsACR - computes the ACR itself
	///
	/// Everything here mirrors what the FAA's own ICAO-ACR program does in
	/// ACRClassDriver/Form1_ICAO.vb. Where a choice existed, the driver's
	/// behaviour was copied rather than reasoned about independently, so that the
	/// output can be cross-checked against the FAA program directly.
	/// </summary>
	public class AcrEngine
	{
		// The library works entirely in US units: libGL is lb, libCP is psi,
		// libTX/libTY are inches (see clsAC.InitACLib). So Metric is always false.
		const bool Metric = false;

		ACClassLib.clsAC.AircraftCharacteristics[] _lib;
		short _count;

		// Resolved once: each ACR evaluation would otherwise re-scan 411 entries,
		// and MaxAllowableWeightLb evaluates many times per aircraft.
		readonly Dictionary<string, int> _indexCache =
			new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

		public string LibraryPath { get; private set; }

		public void Load(string aircraftXmlPath)
		{
			if (!File.Exists(aircraftXmlPath))
				throw new FileNotFoundException(
					"The FAA aircraft library was not found.\r\n\r\nExpected: " + aircraftXmlPath, aircraftXmlPath);

			LibraryPath = aircraftXmlPath;

			ACClassLib.clsAC loader = new ACClassLib.clsAC();
			loader.XMLFileLocation = aircraftXmlPath;

			ACClassLib.clsAC.AircraftCharacteristics[] ac = null;
			short[] groups = null;
			string[] groupNames = null;
			short nac = 0, nbelly = 0, ngroups = 0;

			// InitACLib ReDims every ByRef array itself.
			loader.InitACLib(ref ac, ref groups, ref groupNames, ref nac, ref nbelly, ref ngroups, false);

			_lib = ac;
			_count = nac;
			_indexCache.Clear();
		}

		int IndexOf(string libraryName)
		{
			if (_lib == null)
				throw new InvalidOperationException("The aircraft library has not been loaded.");

			int cached;
			if (_indexCache.TryGetValue(libraryName, out cached))
				return cached;

			for (int i = 0; i <= _count && i < _lib.Length; i++)
			{
				if (string.Equals(_lib[i].libACName, libraryName, StringComparison.OrdinalIgnoreCase))
				{
					_indexCache[libraryName] = i;
					return i;
				}
			}
			throw new InvalidOperationException(
				"Aircraft \"" + libraryName + "\" is not present in this aircraft.xml.");
		}

		/// <summary>Maximum take-off weight from the library, in pounds.</summary>
		public float MaxWeightLb(AircraftSpec spec)
		{
			return _lib[IndexOf(spec.LibraryName)].libGL;
		}

		/// <summary>Tyre (inflation) pressure from the library, in psi.</summary>
		public float TyrePressurePsi(AircraftSpec spec)
		{
			return _lib[IndexOf(spec.LibraryName)].libCP;
		}

		/// <summary>
		/// Copies wheel coordinates 0..n out of the library entry.
		///
		/// libTX/libTY are 1-based: clsAC fills indices 1..Count and leaves slot 0
		/// unused, matching the "index values are 1 through 4" note in the FAA API
		/// document. The driver copies 0..n inclusive, so this does too.
		/// </summary>
		static void CopyCoords(ACClassLib.clsAC.AircraftCharacteristics ac, int wheels,
		                       out float[] x, out float[] y)
		{
			x = new float[wheels + 1];
			y = new float[wheels + 1];

			if (ac.libTX == null || ac.libTY == null)
				throw new InvalidOperationException(
					"Aircraft \"" + ac.libACName + "\" has no wheel coordinates in this aircraft.xml.");

			for (int i = 0; i <= wheels; i++)
			{
				x[i] = i < ac.libTX.Length ? ac.libTX[i] : 0f;
				y[i] = i < ac.libTY.Length ? ac.libTY[i] : 0f;
			}
		}

		/// <summary>
		/// ACR at a given gross weight, for all four subgrade categories.
		/// </summary>
		public AcrResult Acr(AircraftSpec spec, float grossWeightLb, PavementKind pavement)
		{
			ACRClassLib.clsACR.PavementType pt = pavement == PavementKind.Rigid
				? ACRClassLib.clsACR.PavementType.Rigid
				: ACRClassLib.clsACR.PavementType.Flexible;

			if (pavement == PavementKind.Rigid)
			{
				// Rigid ACR uses the single most demanding truck only.
				// Form1_ICAO.vb: wheels_number = libNWheels / 2, percent_gw = libMGpcntPCN.
				//
				// For the 747 the wing and body trucks carry the same share at the
				// same pressure and differ only in wheel spacing, so rather than
				// guessing which is "most demanding" both are evaluated and the
				// higher value is kept for each subgrade independently.
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
			int mainIdx = IndexOf(spec.LibraryName);
			ACClassLib.clsAC.AircraftCharacteristics main = _lib[mainIdx];

			int wheels1 = main.libNWheels;
			float percent1 = main.libMGpcntPCN * 2f;   // the entry covers one side; x2 for the pair
			float[] x1, y1;
			CopyCoords(main, wheels1, out x1, out y1);

			ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
			ACRClassLib.clsACR.ACRdata data;

			if (!spec.HasBodyGear)
			{
				data = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1,
				                           main.libCP, x1, y1, Metric);
			}
			else
			{
				// Wing gear and body gear are supplied as two separate gears, the
				// same way Form1_ICAO.vb handles the 747 and A380.
				ACClassLib.clsAC.AircraftCharacteristics belly = _lib[IndexOf(spec.BellyLibraryName)];

				int wheels2 = belly.libNWheels;
				float percent2 = belly.libMGpcntPCN * 2f;
				float[] x2, y2;
				CopyCoords(belly, wheels2, out x2, out y2);

				data = runner.CalculateACR(pt, grossWeightLb, percent1, wheels1, main.libCP, x1, y1,
				                           percent2, wheels2, belly.libCP, x2, y2, Metric);
			}

			return Unpack(data);
		}

		AcrResult RigidOneTruck(ACRClassLib.clsACR.PavementType pt, string libraryName, float grossWeightLb)
		{
			ACClassLib.clsAC.AircraftCharacteristics ac = _lib[IndexOf(libraryName)];

			int wheels = ac.libNWheels / 2;
			float percent = ac.libMGpcntPCN;
			float[] x, y;
			CopyCoords(ac, wheels, out x, out y);

			ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
			return Unpack(runner.CalculateACR(pt, grossWeightLb, percent, wheels, ac.libCP, x, y, Metric));
		}

		/// <summary>
		/// Reads the ACRdata arrays.
		///
		/// These are length 5 with slot 0 unused, and run backwards:
		/// index 1 is subgrade D and index 4 is subgrade A.
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
		///
		/// ACR rises monotonically with weight, so this is a plain bisection.
		/// Returns MTOW when the aircraft already fits at MTOW, and 0 when it does
		/// not fit even at the lowest weight considered.
		///
		/// Each step runs the layered-elastic solver, which is not cheap, so the
		/// loop stops at 50 lb - far finer than any operational use of the answer.
		/// </summary>
		public float MaxAllowableWeightLb(AircraftSpec spec, PcrCode pcr, out bool limitedByPavement)
		{
			limitedByPavement = false;

			float mtow = MaxWeightLb(spec);
			if (Acr(spec, mtow, pcr.Pavement).For(pcr.Subgrade) <= pcr.Value)
				return mtow;                       // pavement is not the limit

			limitedByPavement = true;

			float low = mtow * 0.20f;              // well below any realistic empty weight
			if (Acr(spec, low, pcr.Pavement).For(pcr.Subgrade) > pcr.Value)
				return 0f;                         // unusable at any sensible weight

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
