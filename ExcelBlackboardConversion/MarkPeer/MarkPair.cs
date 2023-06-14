using System.Collections.Generic;
using System.Linq;

namespace ExcelBlackboardConversion.MarkPeer
{
	internal class MarkPair
	{
		public MarkPair(string refMarker, string deltaMarker)
		{
			RefMarker = refMarker;
			DeltaMarker = deltaMarker;
		}

		public string RefMarker { get; set; }
		public string DeltaMarker { get; set; }
		public List<double> MarkDeltas { get; set; } = new();

		public IEnumerable<double> GetDeltas(string refMarker)
		{
			if (refMarker == RefMarker)
				return MarkDeltas;
			else
				return MarkDeltas.Select(x => x * -1);
		}
	}

	
}
