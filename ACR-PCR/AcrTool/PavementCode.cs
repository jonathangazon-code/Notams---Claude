using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AcrTool
{
	public enum PavementKind
	{
		Flexible = 1,
		Rigid = 2
	}

	/// <summary>Which of the two rating systems a code belongs to.</summary>
	public enum RatingMethod
	{
		/// <summary>ACR/PCR - ICAO Annex 14, in force since 28 Nov 2024.</summary>
		Acr,
		/// <summary>ACN/PCN - the legacy method it replaced.</summary>
		Acn
	}

	/// <summary>
	/// A published PCR or PCN code, e.g. "690/R/B/W/T" or "80/R/B/W/T".
	///
	/// Both systems share the same five-part shape:
	///   1. numerical rating
	///   2. pavement type      R = rigid, F = flexible
	///   3. subgrade category  A, B, C, D
	///   4. tyre pressure      W, X, Y, Z
	///   5. evaluation method  T = technical, U = using aircraft experience
	///
	/// They are NOT interchangeable: the subgrade letters mean different things
	/// (ACR uses subgrade modulus E, ACN uses CBR for flexible and k for rigid),
	/// the numbers are on different scales (ACR is roughly ten times ACN), and the
	/// overload allowance differs on rigid pavements - see OverloadFactor. The tyre
	/// pressure categories, however, are identical in both.
	/// </summary>
	public class PavementCode
	{
		public RatingMethod Method;
		public float Value;
		public PavementKind Pavement;
		public char Subgrade;
		public char TyreCategory;
		public char Evaluation;

		static readonly Regex Pattern = new Regex(
			@"^\s*(?:PCR|PCN)?\s*(\d+(?:\.\d+)?)\s*/\s*([RF])\s*/\s*([ABCD])\s*/\s*([WXYZ])\s*/\s*([TU])\s*$",
			RegexOptions.IgnoreCase);

		/// <summary>
		/// Parses a full code. Returns false with a readable reason rather than
		/// throwing - this is fed straight from a text box.
		/// </summary>
		public static bool TryParse(string text, RatingMethod method, out PavementCode result, out string error)
		{
			result = null;
			error = null;
			string label = method == RatingMethod.Acr ? "PCR" : "PCN";

			if (string.IsNullOrEmpty(text) || text.Trim().Length == 0)
			{
				error = "Enter a " + label + " code, for example "
					+ (method == RatingMethod.Acr ? "690/R/B/W/T" : "80/R/B/W/T");
				return false;
			}

			Match m = Pattern.Match(text);
			if (!m.Success)
			{
				error = "Not a valid " + label + " code. Expected five parts separated by '/', for example "
					+ (method == RatingMethod.Acr ? "690/R/B/W/T" : "80/R/B/W/T")
					+ "  (number / R or F / A-D / W-Z / T or U).";
				return false;
			}

			PavementCode p = new PavementCode();
			p.Method = method;
			p.Value = float.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
			p.Pavement = char.ToUpperInvariant(m.Groups[2].Value[0]) == 'R'
				? PavementKind.Rigid : PavementKind.Flexible;
			p.Subgrade = char.ToUpperInvariant(m.Groups[3].Value[0]);
			p.TyreCategory = char.ToUpperInvariant(m.Groups[4].Value[0]);
			p.Evaluation = char.ToUpperInvariant(m.Groups[5].Value[0]);

			if (p.Value <= 0)
			{
				error = "The numerical " + label + " must be greater than zero.";
				return false;
			}

			result = p;
			return true;
		}

		public override string ToString()
		{
			return string.Format(CultureInfo.InvariantCulture, "{0:0.##}/{1}/{2}/{3}/{4}",
				Value, Pavement == PavementKind.Rigid ? "R" : "F",
				Subgrade, TyreCategory, Evaluation);
		}

		/// <summary>
		/// Maximum tyre pressure the code allows, in psi; float.MaxValue for W.
		///
		/// The same in both systems. ICAO's own before/after comparison of Annex 14
		/// for 28 November 2024 lists an identical table on each side:
		/// W unlimited, X 1.75 MPa, Y 1.25 MPa, Z 0.50 MPa.
		///
		/// An earlier version of this file gave ACN/PCN 1.50 and 1.00 MPa for X and
		/// Y. Those are the pre-2008 categories, revised by ICAO long before the
		/// move to ACR/PCR, and wrong for any PCN published today - they would have
		/// refused aircraft that the pavement actually accepts.
		/// </summary>
		public float TyrePressureLimitPsi()
		{
			switch (TyreCategory)
			{
				case 'W': return float.MaxValue;   // no limit
				case 'X': return 254f;             // 1.75 MPa
				case 'Y': return 181f;             // 1.25 MPa
				case 'Z': return 73f;              // 0.50 MPa
				default:  return float.MaxValue;
			}
		}

		/// <summary>
		/// Overload tolerance factor.
		///
		///   ACR/PCR   +10% on flexible AND rigid
		///   ACN/PCN   +10% on flexible, +5% on rigid or composite
		///
		/// The rigid allowance was raised from 5% to 10% with the move to ACR/PCR.
		/// Annex 14 now reads "for flexible and rigid pavements, occasional
		/// movements by aircraft with ACR not exceeding 10 per cent above the
		/// reported PCR should not adversely affect the pavement". The legacy
		/// method kept the two apart, so the PCN half must not inherit the ACR
		/// figure.
		///
		/// Both carry conditions this tool cannot check: overload movements should
		/// stay at roughly 5% of annual movements at most, none are acceptable on a
		/// pavement showing distress or failure or during a thaw, and anything past
		/// the allowance needs a cumulative-damage analysis.
		///
		/// Whether to allow any of it is the aerodrome operator's decision, so this
		/// is never applied unless the dispatcher explicitly ticks the box.
		/// </summary>
		public float OverloadFactor()
		{
			if (Method == RatingMethod.Acr) return 1.10f;
			return Pavement == PavementKind.Flexible ? 1.10f : 1.05f;
		}

		/// <summary>The limit actually compared against, tolerance included or not.</summary>
		public float EffectiveValue(bool allowOverload)
		{
			return allowOverload ? Value * OverloadFactor() : Value;
		}

		/// <summary>e.g. "+10%" - for labelling what was applied.</summary>
		public string OverloadText()
		{
			return OverloadFactor() >= 1.0999f ? "+10%" : "+5%";
		}

		public string TyreCategoryText()
		{
			switch (TyreCategory)
			{
				case 'W': return "W - unlimited";
				case 'X': return "X - high, max 1.75 MPa (254 psi)";
				case 'Y': return "Y - medium, max 1.25 MPa (181 psi)";
				case 'Z': return "Z - low, max 0.50 MPa (73 psi)";
				default:  return TyreCategory.ToString();
			}
		}

		public string EvaluationText()
		{
			return Evaluation == 'T' ? "T - technical evaluation" : "U - using aircraft experience";
		}

		public string SubgradeText()
		{
			if (Method == RatingMethod.Acr)
				return "subgrade " + Subgrade;

			// Legacy ACN subgrade categories are defined by CBR (flexible) or k (rigid).
			if (Pavement == PavementKind.Flexible)
			{
				switch (Subgrade)
				{
					case 'A': return "subgrade A (CBR 15, high)";
					case 'B': return "subgrade B (CBR 10, medium)";
					case 'C': return "subgrade C (CBR 6, low)";
					case 'D': return "subgrade D (CBR 3, ultra low)";
				}
			}
			else
			{
				switch (Subgrade)
				{
					case 'A': return "subgrade A (k = 150 MN/m3, high)";
					case 'B': return "subgrade B (k = 80 MN/m3, medium)";
					case 'C': return "subgrade C (k = 40 MN/m3, low)";
					case 'D': return "subgrade D (k = 20 MN/m3, ultra low)";
				}
			}
			return "subgrade " + Subgrade;
		}
	}
}
