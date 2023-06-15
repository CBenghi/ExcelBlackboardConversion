using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion.MarkPeer
{
	internal class SingleMark
	{
		public string Marker { get; set; } = string.Empty;
		public double Mark { get; set; } = 0;
	}

	internal class PairingCollection : IEnumerable<MarkPair> 
	{
		public IEnumerable<string> GetAllMarkers()
		{
			return Pairs.Select(x => x.RefMarker)
				.Concat(Pairs.Select(x => x.DeltaMarker))
				.Distinct().OrderBy(x=> x);
		}

		public IEnumerable<double> GetDeltas(string marker)
		{
			List<double> result = new List<double>();	
			foreach (var pair in Pairs.Where(x=>x.RefMarker == marker)) 
			{
				result.AddRange(pair.GetDeltas(marker));
			}
			return result;
		}

		private List<MarkPair> Pairs { get; set; } = new();

		public IEnumerator<MarkPair> GetEnumerator()
		{
			return ((IEnumerable<MarkPair>)Pairs).GetEnumerator();
		}

		public bool TryGet(string m1, string m2, [NotNullWhen(true)] out MarkPair? pair)
		{
			var t = Pairs.FirstOrDefault(x => x.RefMarker == m1 && x.DeltaMarker == m2);
			if (t != null)
			{
				pair = t;
				return true;
			}
				
			t = Pairs.FirstOrDefault(x => x.RefMarker == m2 && x.DeltaMarker == m1);
			if (t != null)
			{
				pair = t;
				return true;
			}
			pair = null;
			return false;
		}

		internal MarkPair GetOrAdd(string m1, string m2)
		{
			if (TryGet(m1, m2, out var t))
				return t;
			t = new MarkPair(m1, m2);
			Pairs.Add(t);
			return t;
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return ((IEnumerable)Pairs).GetEnumerator();
		}

		
	}

	
}
