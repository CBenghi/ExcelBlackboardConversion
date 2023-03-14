using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion
{
    internal class Student
    {
        public string Id { get; set; }
        public int ExcelRow { get; set; }

        public Dictionary<string, Answer> Answers { get; set; } = new();
        public string Cohort { get; internal set; }
        public string ComputedMark { get; internal set; }
        public string Mark { get; internal set; }
    }
}
