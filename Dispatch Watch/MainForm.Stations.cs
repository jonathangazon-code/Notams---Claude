using System.Collections.Generic;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Cache ICAO → [IATA, LH, FedEx, Charters] — chargé au démarrage, rafraîchi après modification des stations.
		private static Dictionary<string, string[]> _stationsCache = null;

		public static void LoadStationsCache()
		{
			_stationsCache = new Dictionary<string, string[]>(System.StringComparer.OrdinalIgnoreCase);
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA", conn).ExecuteReader();
			while (reader.Read())
			{
				string icao = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				if (icao == "") continue;
				_stationsCache[icao] = new string[]
				{
					!reader.IsDBNull(2) ? reader.GetString(2) : "",
					!reader.IsDBNull(3) ? reader.GetString(3) : "",
					!reader.IsDBNull(4) ? reader.GetString(4) : "",
					!reader.IsDBNull(5) ? reader.GetString(5) : ""
				};
			}
			conn.Close();
		}

		public static string IsOpsType(string OpsType, string location)
		{
			if (_stationsCache == null) LoadStationsCache();
			string[] row;
			if (!_stationsCache.TryGetValue(location, out row)) return "";
			if (OpsType == "LH")       return row[1];
			if (OpsType == "FedEx")    return row[2];
			if (OpsType == "Charters") return row[3];
			return "";
		}

		public static string GetIATA(string location)
		{
			if (_stationsCache == null) LoadStationsCache();
			string[] row;
			if (!_stationsCache.TryGetValue(location, out row)) return "";
			return row[0];
		}
	}
}
