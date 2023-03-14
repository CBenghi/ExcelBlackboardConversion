using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion
{
    internal class Answer
    {
        public string QuestionId { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public double AutoScore { get; set; } = -1;
        public double ManualScore { get; set; } = -1;

        internal double GetScore()
        {
            if (ManualScore != -1)
                return ManualScore;
            if (AutoScore != -1) 
                return AutoScore;
            return 0;
        }

        internal bool IsManual()
        {
            return ManualScore != -1;
        }
    }
}
