namespace ICAO_CSV
{
	public partial class MainForm
	{
		private static readonly string[] _monthAbbrevs = { "", "JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC" };

		public static string MonthAbbrev(string mm)
		{
			int m;
			if (int.TryParse(mm, out m) && m >= 1 && m <= 12) return _monthAbbrevs[m];
			return mm;
		}

		public static string dateTransformation(string notamDate)
		{
			notamDate = notamDate.Substring(0, 16);
			return notamDate.Substring(8, 2) + MonthAbbrev(notamDate.Substring(5, 2)) + notamDate.Substring(0, 4);
		}
	}
}
