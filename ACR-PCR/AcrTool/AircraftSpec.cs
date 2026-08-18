using System;
using System.Collections.Generic;

namespace AcrTool
{
	/// <summary>
	/// One aircraft the tool reports on, mapped onto entries of the FAA aircraft
	/// library (aircraft.xml).
	///
	/// The freighters have no entry of their own: -400F / -400ERF share landing
	/// gear geometry and tyre pressure with the passenger -400 / -400ER, and the
	/// library's gross weights are already the freighter figures (877 000 lb and
	/// 913 000 lb). LibraryName is shown in the UI so the substitution is visible
	/// and checkable rather than silently assumed.
	/// </summary>
	public class AircraftSpec
	{
		/// <summary>Label shown to the dispatcher.</summary>
		public string Display;

		/// <summary>Name of the main entry in aircraft.xml.</summary>
		public string LibraryName;

		/// <summary>
		/// Name of the companion body ("Belly") entry, or null for a
		/// conventional two-truck aircraft.
		///
		/// The 747 has four main trucks: two under the wings and two under the
		/// body. The FAA library splits them into two entries, and the flexible
		/// ACR must account for both - see AcrEngine.Acr.
		/// </summary>
		public string BellyLibraryName;

		// No per-type weight-share override is applied. FF_sub.vb's SetPCN_for_AC
		// looks like it overrides one (93.32/100/4 for the 747-400), but it writes
		// libMGpcnt while every ACR path reads libMGpcntPCN, and the line that
		// would have written libMGpcntPCN is commented out. The library already
		// carries the same figure: aircraft.xml gives MgPercentPCN = 0.2333 for
		// the 747-400, which is exactly 93.32/100/4.

		public static List<AircraftSpec> Fleet()
		{
			List<AircraftSpec> list = new List<AircraftSpec>();

			list.Add(new AircraftSpec {
				Display = "B737-400",
				LibraryName = "B737-400"
			});

			list.Add(new AircraftSpec {
				Display = "B737-800",
				LibraryName = "B737-800"
			});

			list.Add(new AircraftSpec {
				Display = "B747-400F",
				LibraryName = "B747-400",
				BellyLibraryName = "B747-400 Belly"
			});

			list.Add(new AircraftSpec {
				Display = "B747-400ERF",
				LibraryName = "B747-400ER",
				BellyLibraryName = "B747-400ER Belly"
			});

			return list;
		}

		public bool HasBodyGear
		{
			get { return !string.IsNullOrEmpty(BellyLibraryName); }
		}
	}
}
