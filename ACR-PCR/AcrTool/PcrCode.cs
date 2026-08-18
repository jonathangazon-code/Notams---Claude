using System;
using System.Text.RegularExpressions;

namespace AcrTool
{
	/// <summary>
	/// A published PCR code, e.g. "690/R/B/W/T".
	///
	/// Five components, per ICAO Annex 14 (ACR-PCR, in force 28 Nov 2024):
	///   1. numerical PCR
	///   2. pavement type      R = rigid, F = flexible
	///   3. subgrade category  A, B, C, D
	///   4. tyre pressure      W = unlimited, X = high, Y = medium, Z = low
	///   5. evaluation method  T = technical, U = using aircraft experience
	/// </summary>
	public class PcrCode
	{
		public float Value;
		public PavementKind Pavement;
		public char Subgrade;
		public char TyreCategory;
		public char Method;

		// Allows extra spaces and lower case; the separator is always '/'.
		static readonly Regex Pattern = new Regex(
			@"^\s*(\d+(?:\.\d+)?)\s*/\s*([RF])\s*/\s*([ABCD])\s*/\s*([WXYZ])\s*/\s*([TU])\s*$",
			RegexOptions.IgnoreCase);

		/// <summary>
		/// Parses a full PCR code. Returns false with a human-readable reason
		/// rather than throwing - this is fed directly by a text box.
		/// </summary>
		public static bool TryParse(string text, out PcrCode result, out string error)
		{
			result = null;
			error = null;

			if (string.IsNullOrEmpty(text) || text.Trim().Length == 0)
			{
				error = "Enter a PCR code, for example 690/R/B/W/T";
				return false;
			}

			Match m = Pattern.Match(text);
			if (!m.Success)
			{
				error = "Not a valid PCR code. Expected five parts separated by '/', "
					+ "for example 690/R/B/W/T  (number / R or F / A-D / W-Z / T or U).";
				return false;
			}

			PcrCode p = new PcrCode();
			p.Value = float.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
			p.Pavement = char.ToUpperInvariant(m.Groups[2].Value[0]) == 'R'
				? PavementKind.Rigid : PavementKind.Flexible;
			p.Subgrade = char.ToUpperInvariant(m.Groups[3].Value[0]);
			p.TyreCategory = char.ToUpperInvariant(m.Groups[4].Value[0]);
			p.Method = char.ToUpperInvariant(m.Groups[5].Value[0]);

			if (p.Value <= 0)
			{
				error = "The numerical PCR must be greater than zero.";
				return false;
			}

			result = p;
			return true;
		}

		public override string ToString()
		{
			return string.Format(System.Globalization.CultureInfo.InvariantCulture,
				"{0:0.##}/{1}/{2}/{3}/{4}",
				Value, Pavement == PavementKind.Rigid ? "R" : "F",
				Subgrade, TyreCategory, Method);
		}

		/// <summary>
		/// Maximum tyre pressure allowed by the code letter, in psi.
		/// float.MaxValue for W (no limit). Source: ICAO Annex 14 tyre pressure categories.
		/// </summary>
		public float TyrePressureLimitPsi()
		{
			switch (TyreCategory)
			{
				case 'W': return float.MaxValue;   // unlimited
				case 'X': return 254.0f;           // high,   <= 1.75 MPa
				case 'Y': return 181.0f;           // medium, <= 1.25 MPa
				case 'Z': return 73.0f;            // low,    <= 0.50 MPa
				default:  return float.MaxValue;
			}
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

		public string MethodText()
		{
			return Method == 'T' ? "T - technical evaluation" : "U - using aircraft experience";
		}
	}

	public enum PavementKind
	{
		Flexible = 1,
		Rigid = 2
	}
}
