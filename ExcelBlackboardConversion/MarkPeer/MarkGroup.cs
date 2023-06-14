using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelBlackboardConversion.MarkPeer
{
	internal class MarkGroup
	{
		public List<SingleMark> Marks { get; set; } = new();

		public void PopulateCollection(ref PairingCollection pairingCollection)
		{
			for (int i = 0; i < Marks.Count - 1; i++)
			{
				for (int j = i + 1; j < Marks.Count; j++)
				{
					var m1 = Marks[i].Marker;
					var m2 = Marks[j].Marker;
					var p = pairingCollection.GetOrAdd(m1, m2);
					var delta = Marks[j].Mark - Marks[i].Mark;
					if (p.RefMarker == m2) 
						delta *= -1;
					p.MarkDeltas.Add(delta);
				}
			}
		}

		public static IEnumerable<MarkGroup> GetGroups(string filename)
		{
			FileInfo f = new FileInfo(filename);
			using var package = new ExcelPackage(f);
			// prepare question dictionary
			var table = package.Workbook.Worksheets.FirstOrDefault();
			if (table is null)
				yield break;
			var magicNumber = table.Cells["B1"].Text;
			if (magicNumber is null || magicNumber != "Marker Name")
				yield break;

			magicNumber = table.Cells["K1"].Text;
			if (magicNumber is null || magicNumber != "Total")
				yield break;
			int row = 2;
			var g = new MarkGroup();
			while (true)
			{
				var name = table.Cells[$"B{row}"].Text;
				var mark = table.Cells[$"K{row}"].Text;
				if (string.IsNullOrEmpty(name) && !g.Marks.Any())
					yield break;
				if (name is null || name == string.Empty)
				{
					yield return g;
					g = new MarkGroup();
				}
				if (!string.IsNullOrEmpty(mark))
				{
					var mk = double.Parse(mark);
					g.Marks.Add(new SingleMark() { Marker = name, Mark = mk });
				}
				row++;
			}
		}
	}

	
}
