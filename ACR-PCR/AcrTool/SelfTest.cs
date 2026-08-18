using System;
using System.Text;

namespace AcrTool
{
	/// <summary>
	/// Replays the worked example printed in the FAA API document
	/// ("User Information for ICAO-ACR", section 2.4, Example 1).
	///
	/// This is the one check that matters. It exercises the interop end to end -
	/// argument order, US units, 1-based coordinate arrays and the reversed
	/// subgrade indexing - against numbers published by the FAA. If it fails,
	/// nothing else the tool reports can be trusted.
	/// </summary>
	public static class SelfTest
	{
		// Generic dual-tandem (2D-400) gear on rigid pavement, US units.
		const float GrossWeightLb = 400000f;
		const float PercentGw = 0.475f;
		const int Wheels = 4;
		const float TyrePressurePsi = 200f;

		// libACR(1) through libACR(4), i.e. subgrade D, C, B, A.
		static readonly float[] Expected = { 0f, 894.672058f, 817.8752f, 744.0264f, 641.6398f };

		/// <summary>Runs the reference case. Returns true when every value matches.</summary>
		public static bool Run(out string report)
		{
			StringBuilder sb = new StringBuilder();
			bool ok = true;

			try
			{
				// Declared in the document as X1(4)/Y1(4): five slots, filled 1..4.
				float[] x = new float[5];
				float[] y = new float[5];
				x[1] = -15f; y[1] = 0f;
				x[2] = 15f;  y[2] = 0f;
				x[3] = -15f; y[3] = 55f;
				x[4] = 15f;  y[4] = 55f;

				ACRClassLib.clsACR runner = new ACRClassLib.clsACR();
				ACRClassLib.clsACR.ACRdata data = runner.CalculateACR(
					ACRClassLib.clsACR.PavementType.Rigid,
					GrossWeightLb, PercentGw, Wheels, TyrePressurePsi,
					x, y, false);

				sb.AppendLine("FAA reference case - 2D-400 gear, rigid pavement, US units");
				sb.AppendLine("400 000 lb, 47.5% on the gear, 4 wheels, 200 psi");
				sb.AppendLine();
				sb.AppendLine("Subgrade   computed        expected        delta");

				string[] names = { "", "D", "C", "B", "A" };
				for (int i = 1; i <= 4; i++)
				{
					float got = data.libACR[i];
					float want = Expected[i];
					float delta = got - want;

					// The library returns Single; anything beyond rounding noise is a real failure.
					bool pass = Math.Abs(delta) <= 0.05f;
					if (!pass) ok = false;

					sb.AppendLine(string.Format(
						"   {0}       {1,10:0.0000}      {2,10:0.0000}     {3,8:+0.0000;-0.0000; 0.0000}   {4}",
						names[i], got, want, delta, pass ? "ok" : "FAILED"));
				}

				sb.AppendLine();
				sb.AppendLine(ok
					? "PASS - the calculation engine is wired up correctly."
					: "FAIL - do not use the results of this tool until this is resolved.");
			}
			catch (Exception ex)
			{
				ok = false;
				sb.AppendLine("The self-test could not run.");
				sb.AppendLine();
				sb.AppendLine(ex.GetType().Name + ": " + ex.Message);
				sb.AppendLine();
				sb.AppendLine("Most likely ACRClassLib.dll is missing from the folder "
					+ "containing this program, or it targets a newer .NET than the one installed.");
			}

			report = sb.ToString();
			return ok;
		}
	}
}
