using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion
{
    internal class Question
    {
        public string ID { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public int StartingColumn { get; set; }
        public double PossiblePoints { get; set; }

        public enum InformationColumn
        {
            Answer,
            AutoScore,
            ManualScore
        }

        internal int GetColumn(InformationColumn column)
        {
            switch (column)
            {
                case InformationColumn.Answer:
                    return StartingColumn + 2;
                    
                case InformationColumn.AutoScore:
                    return StartingColumn + 4;
                case InformationColumn.ManualScore:
                    return StartingColumn + 5;
                default:
                    return 0;
            }
        }
    }
}
