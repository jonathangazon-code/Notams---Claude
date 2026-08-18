using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace AcrTool
{
	/// <summary>One published ACN point: a weight and the eight ACN values at it.</summary>
	public class AcnPoint
	{
		public float WeightLb;
		public float FlexA, FlexB, FlexC, FlexD;
		public float RigidA, RigidB, RigidC, RigidD;

		public float Value(PavementKind pavement, char subgrade)
		{
			bool rigid = pavement == PavementKind.Rigid;
			switch (char.ToUpperInvariant(subgrade))
			{
				case 'A': return rigid ? RigidA : FlexA;
				case 'B': return rigid ? RigidB : FlexB;
				case 'C': return rigid ? RigidC : FlexC;
				case 'D': return rigid ? RigidD : FlexD;
				default: throw new ArgumentException("Unknown subgrade category: " + subgrade);
			}
		}
	}

	public class AcnAircraft
	{
		public string Name;
		public float TyrePressurePsi;
		/// <summary>Published points, sorted by weight ascending. At least two.</summary>
		public List<AcnPoint> Points = new List<AcnPoint>();
	}

	/// <summary>
	/// ACN from published tables rather than from a calculation.
	///
	/// There is no callable FAA library for the legacy method: ICAO-ACN 1.0 ships
	/// only a GUI program, with no API, no source and no documented entry point
	/// (see vendor/faa-acn/README.txt). Re-implementing the method - CBR with the
	/// 2007 alpha factors for flexible, PCA for rigid - would be a large piece of
	/// numerical work that could not be verified from here.
	///
	/// Manufacturers publish ACN directly, so those values are used as given, and
	/// intermediate weights are interpolated. ICAO publishes ACN at maximum and
	/// empty weights precisely so it can be interpolated between them: ACN varies
	/// essentially linearly with weight. Extra intermediate points are honoured if
	/// the data file provides them, giving a piecewise-linear curve.
	/// </summary>
	public class AcnEngine : IRatingEngine
	{
		Dictionary<string, AcnAircraft> _byName =
			new Dictionary<string, AcnAircraft>(StringComparer.OrdinalIgnoreCase);

		string _source = "";
		string _error;

		public string RatingName { get { return "ACN"; } }
		public RatingMethod Method { get { return RatingMethod.Acn; } }
		public bool Ready { get { return _error == null && _byName.Count > 0; } }
		public string NotReadyReason { get { return Ready ? null : _error ?? "No ACN table data has been entered yet."; } }

		public string Provenance
		{
			get
			{
				return Ready
					? "ACN read from published tables, interpolated by weight. Source: " + _source
					: "ACN table not loaded - " + NotReadyReason;
			}
		}

		public void Load(string path)
		{
			_byName.Clear();
			_error = null;

			try
			{
				if (!File.Exists(path))
				{
					_error = "acn-data.xml was not found next to the program.";
					return;
				}

				XDocument doc = XDocument.Load(path);
				XElement root = doc.Root;
				XAttribute src = root.Attributes().FirstOrDefault(a => a.Name.LocalName == "source");
				_source = src == null ? "(not stated in acn-data.xml)" : src.Value;

				foreach (XElement ac in root.Elements().Where(e => e.Name.LocalName == "Aircraft"))
				{
					AcnAircraft a = new AcnAircraft();
					a.Name = Attr(ac, "name");
					a.TyrePressurePsi = Num(Attr(ac, "tyrePressurePsi"));

					foreach (XElement p in ac.Elements().Where(e => e.Name.LocalName == "Point"))
					{
						AcnPoint pt = new AcnPoint();

						// Published tables are usually in kg; either attribute is
						// accepted so the file can keep the source's own unit.
						string kg = Attr(p, "weightKg");
						pt.WeightLb = string.IsNullOrEmpty(kg)
							? Num(Attr(p, "weightLb"))
							: Num(kg) * AcrEngine.LbPerKg;
						pt.FlexA = Num(Attr(p, "flexA"));  pt.RigidA = Num(Attr(p, "rigidA"));
						pt.FlexB = Num(Attr(p, "flexB"));  pt.RigidB = Num(Attr(p, "rigidB"));
						pt.FlexC = Num(Attr(p, "flexC"));  pt.RigidC = Num(Attr(p, "rigidC"));
						pt.FlexD = Num(Attr(p, "flexD"));  pt.RigidD = Num(Attr(p, "rigidD"));

						// A placeholder row with nothing filled in is not data.
						if (pt.WeightLb > 0) a.Points.Add(pt);
					}

					a.Points = a.Points.OrderBy(x => x.WeightLb).ToList();
					if (!string.IsNullOrEmpty(a.Name) && a.Points.Count >= 2)
						_byName[a.Name] = a;
				}

				if (_byName.Count == 0)
					_error = "acn-data.xml has no aircraft with at least two published weight points. "
						+ "Fill in the ACN values from the manufacturer tables.";
			}
			catch (Exception ex)
			{
				_error = "acn-data.xml could not be read: " + ex.Message;
			}
		}

		AcnAircraft Get(AircraftSpec spec)
		{
			AcnAircraft a;
			if (!_byName.TryGetValue(spec.Display, out a))
				throw new InvalidOperationException(
					"No ACN table for \"" + spec.Display + "\" in acn-data.xml.");
			return a;
		}

		public float MaxWeightLb(AircraftSpec spec)
		{
			AcnAircraft a = Get(spec);
			return a.Points[a.Points.Count - 1].WeightLb;
		}

		/// <summary>Lowest published weight - the empty-weight point of the table.</summary>
		public float MinWeightLb(AircraftSpec spec)
		{
			return Get(spec).Points[0].WeightLb;
		}

		public float TyrePressurePsi(AircraftSpec spec)
		{
			return Get(spec).TyrePressurePsi;
		}

		/// <summary>
		/// ACN at a weight, linearly interpolated between the two published points
		/// that bracket it. Outside the published range the nearest point is used
		/// rather than extrapolating - published tables do not support that.
		/// </summary>
		public float Rating(AircraftSpec spec, float weightLb, PavementCode code)
		{
			AcnAircraft a = Get(spec);
			List<AcnPoint> pts = a.Points;

			if (weightLb <= pts[0].WeightLb)
				return pts[0].Value(code.Pavement, code.Subgrade);
			if (weightLb >= pts[pts.Count - 1].WeightLb)
				return pts[pts.Count - 1].Value(code.Pavement, code.Subgrade);

			for (int i = 1; i < pts.Count; i++)
			{
				if (weightLb <= pts[i].WeightLb)
				{
					AcnPoint lo = pts[i - 1], hi = pts[i];
					float span = hi.WeightLb - lo.WeightLb;
					if (span <= 0) return hi.Value(code.Pavement, code.Subgrade);

					float t = (weightLb - lo.WeightLb) / span;
					float a0 = lo.Value(code.Pavement, code.Subgrade);
					float a1 = hi.Value(code.Pavement, code.Subgrade);
					return a0 + t * (a1 - a0);
				}
			}
			return pts[pts.Count - 1].Value(code.Pavement, code.Subgrade);
		}

		/// <summary>
		/// Heaviest weight whose ACN still fits the PCN.
		///
		/// The curve is piecewise linear, so this inverts it directly segment by
		/// segment - exact, and no search needed (unlike the ACR side, where the
		/// engine is a black box and the answer has to be bisected).
		/// </summary>
		public float MaxAllowableWeightLb(AircraftSpec spec, PavementCode code, float limit, out bool limitedByPavement)
		{
			limitedByPavement = false;

			AcnAircraft a = Get(spec);
			List<AcnPoint> pts = a.Points;

			float top = pts[pts.Count - 1].WeightLb;
			if (Rating(spec, top, code) <= limit)
				return top;                                  // pavement is not the limit

			limitedByPavement = true;

			float bottom = pts[0].WeightLb;
			if (Rating(spec, bottom, code) > limit)
				return 0f;                                   // does not fit even at the lightest published weight

			// Walk down from the top and find the segment where the curve crosses.
			for (int i = pts.Count - 1; i >= 1; i--)
			{
				AcnPoint lo = pts[i - 1], hi = pts[i];
				float aLo = lo.Value(code.Pavement, code.Subgrade);
				float aHi = hi.Value(code.Pavement, code.Subgrade);

				if (aLo <= limit && limit <= aHi)
				{
					float da = aHi - aLo;
					if (da <= 0) return hi.WeightLb;
					float t = (limit - aLo) / da;
					return lo.WeightLb + t * (hi.WeightLb - lo.WeightLb);
				}
			}
			return bottom;
		}

		static string Attr(XElement e, string name)
		{
			XAttribute a = e.Attributes().FirstOrDefault(x => x.Name.LocalName == name);
			return a == null ? null : a.Value;
		}

		static float Num(string text)
		{
			if (string.IsNullOrEmpty(text)) return 0f;
			float v;
			return float.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0f;
		}
	}
}
