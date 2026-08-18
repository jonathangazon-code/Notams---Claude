using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace AcrTool
{
	/// <summary>
	/// One aircraft as published in the FAA aircraft library (aircraft.xml).
	/// Everything is in US units, which is what the ACR engine expects.
	/// </summary>
	public class AircraftEntry
	{
		public string Name;
		public float GrossWeightLb;      // _GrossWeight/us
		public float TyrePressurePsi;    // Cp/us
		public float MainGearPercent;    // MgPercentPCN, fraction of gross weight on one truck pair
		public float[] X;                // wheel coordinates, inches, 1-based (slot 0 unused)
		public float[] Y;

		/// <summary>Number of wheel coordinates published for this entry.</summary>
		public int WheelCount { get { return X.Length - 1; } }
	}

	/// <summary>
	/// Reads aircraft.xml directly.
	///
	/// This replaces the FAA's own ACClassLib/ICAOModels projects, which cannot be
	/// built here: they are ToolsVersion 15.0 (VS2017) and use VB14 syntax that
	/// SharpDevelop's compiler does not accept. Nothing is lost by dropping them -
	/// the five values below are all this tool ever used, and clsAC.InitACLib
	/// reads every one of them straight out of the same XML with no transformation
	/// (clsAC.vb lines 222-274):
	///
	///     libCP          = Cp.UsCustomary
	///     libGL          = GrossWeight.UsCustomary
	///     libMGpcntPCN   = MgPercentPCN
	///     libNWheels     = WheelCoordinates.Count
	///     libTX/libTY    = WheelCoordinates(j).X/Y.UsCustomary, written to index j+1
	///
	/// The derived geometry in modAC.vb (evaluation points, gear-type special
	/// cases, wheel tracks) is never used by this tool, so none of it is ported.
	///
	/// Elements are matched on local name only: the file carries two namespaces
	/// and a fully qualified lookup silently finds nothing - the same trap already
	/// documented for the briefing feed in Dispatch Watch.
	/// </summary>
	public static class AircraftLibrary
	{
		public static Dictionary<string, AircraftEntry> Load(string path)
		{
			if (!File.Exists(path))
				throw new FileNotFoundException(
					"The FAA aircraft library was not found." + Environment.NewLine +
					Environment.NewLine + "Expected: " + path, path);

			XDocument doc = XDocument.Load(path);

			Dictionary<string, AircraftEntry> byName =
				new Dictionary<string, AircraftEntry>(StringComparer.OrdinalIgnoreCase);

			foreach (XElement plane in doc.Descendants()
				.Where(e => e.Attributes().Any(a => a.Name.LocalName == "type" && a.Value.EndsWith("AirplaneInfo"))))
			{
				string name = ChildText(plane, "Name");
				if (string.IsNullOrEmpty(name) || byName.ContainsKey(name)) continue;

				XElement coords = Child(plane, "WheelCoordinates");
				if (coords == null) continue;

				List<XElement> wheels = coords.Elements()
					.Where(e => e.Attributes().Any(a => a.Name.LocalName == "type"
					                                 && a.Value.EndsWith("LengthCoordinates")))
					.ToList();
				if (wheels.Count == 0) continue;

				AircraftEntry entry = new AircraftEntry();
				entry.Name = name;
				entry.GrossWeightLb = UsValue(plane, "_GrossWeight");
				entry.TyrePressurePsi = UsValue(plane, "Cp");
				entry.MainGearPercent = Number(ChildText(plane, "MgPercentPCN"));

				// 1-based, matching clsAC: index 0 stays zero and is never a wheel.
				entry.X = new float[wheels.Count + 1];
				entry.Y = new float[wheels.Count + 1];
				for (int j = 0; j < wheels.Count; j++)
				{
					entry.X[j + 1] = UsValue(wheels[j], "X");
					entry.Y[j + 1] = UsValue(wheels[j], "Y");
				}

				byName[name] = entry;
			}

			if (byName.Count == 0)
				throw new InvalidDataException(
					"No aircraft could be read from " + path + "." + Environment.NewLine +
					"The file does not look like an FAA aircraft library.");

			return byName;
		}

		/// <summary>LibraryVersion attribute from the root element, for provenance.</summary>
		public static string Version(string path)
		{
			try
			{
				XDocument doc = XDocument.Load(path);
				XAttribute a = doc.Root.Attributes().FirstOrDefault(x => x.Name.LocalName == "LibraryVersion");
				return a == null ? "unknown" : a.Value;
			}
			catch { return "unknown"; }
		}

		static XElement Child(XElement parent, string localName)
		{
			return parent.Elements().FirstOrDefault(e => e.Name.LocalName == localName);
		}

		static string ChildText(XElement parent, string localName)
		{
			XElement e = Child(parent, localName);
			return e == null ? null : e.Value;
		}

		/// <summary>
		/// Reads the US customary half of a dimensioned value, e.g.
		/// &lt;Cp&gt;&lt;si&gt;1275.53&lt;/si&gt;&lt;us&gt;185&lt;/us&gt;&lt;/Cp&gt;.
		/// </summary>
		static float UsValue(XElement parent, string localName)
		{
			XElement e = Child(parent, localName);
			if (e == null) return 0f;
			XElement us = Child(e, "us");
			return Number(us == null ? e.Value : us.Value);
		}

		static float Number(string text)
		{
			if (string.IsNullOrEmpty(text)) return 0f;
			float v;
			// The file is machine-written and always uses '.' regardless of locale.
			return float.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0f;
		}
	}
}
