namespace AcrTool
{
	/// <summary>
	/// What the results panel needs from either rating method, so that ACR and ACN
	/// are rendered by exactly the same code and cannot drift apart.
	/// </summary>
	public interface IRatingEngine
	{
		/// <summary>"ACR" or "ACN" - used in column headers and messages.</summary>
		string RatingName { get; }

		RatingMethod Method { get; }

		/// <summary>False when the underlying data could not be loaded.</summary>
		bool Ready { get; }

		/// <summary>Why it is not ready, for display. Null when Ready.</summary>
		string NotReadyReason { get; }

		/// <summary>Short line naming where the numbers come from.</summary>
		string Provenance { get; }

		/// <summary>Maximum weight from the source data, in pounds.</summary>
		float MaxWeightLb(AircraftSpec spec);

		/// <summary>Tyre pressure, psi.</summary>
		float TyrePressurePsi(AircraftSpec spec);

		/// <summary>The rating (ACR or ACN) at a given weight, for the code's subgrade.</summary>
		float Rating(AircraftSpec spec, float weightLb, PavementCode code);

		/// <summary>
		/// Heaviest weight still accepted, capped at maximum weight.
		///
		/// <paramref name="limit"/> is passed in rather than read from the code,
		/// so that any overload tolerance is decided once, by the caller, and is
		/// never silently baked in here.
		///
		/// limitedByPavement is false when the pavement is not the binding limit;
		/// returns 0 when the aircraft does not fit at any sensible weight.
		/// </summary>
		float MaxAllowableWeightLb(AircraftSpec spec, PavementCode code, float limit, out bool limitedByPavement);
	}
}
