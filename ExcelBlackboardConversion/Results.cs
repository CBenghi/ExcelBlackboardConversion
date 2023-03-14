using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion
{
    internal class Results
    {
        public Dictionary<string, Question> Questions { get; set; } = new();
        public Dictionary<string, Student> Students { get; set; } = new();

        internal static Results? FromFile(FileInfo file)
        {
            var ret = new Results();
            using (var package = new ExcelPackage(file))
            {
                // prepare question dictionary
                var table = package.Workbook.Worksheets.FirstOrDefault();
                if (table is null)
                    return null;
                var magicNumber = table.Cells["A1"].Text;
                if (magicNumber is null || magicNumber != "Student ID")    
                    return null;

                int iQcol = 2;
                int iColComment = -1;
                int iColComputed = -1;
                int iColMark = -1;
                while (true) // getting questions and table structure
                {
                    var header = GetAddress(iQcol, 1);
                    var headContent = table.Cells[header].Text;
                    
                    if (headContent.StartsWith("Question ID"))
                    {
                        var qId = table.Cells[GetAddress(iQcol, 2)].Text;
                        var qText = table.Cells[GetAddress(iQcol + 1, 2)].Text;
                        var qPts = table.Cells[GetAddress(iQcol + 3, 2)].Text;
                        Question q = new Question()
                        {
                            ID = qId,
                            PossiblePoints = Convert.ToDouble(qPts),
                            StartingColumn = iQcol,
                            Text = qText
                        };
                        ret.Questions.Add(qId, q);
                        iQcol += 5;
                    }
                    else if (headContent == "")
                        break;
                    if (headContent == "Comment") 
                        iColComment = iQcol;
                    else if (headContent == "Overall Mark") 
                        iColComputed = iQcol;
                    else if (headContent == "Module Mark") 
                        iColMark = iQcol;
                    iQcol++;
                }


                var studentRow = 2;
                while (true)
                {
                    // initialize student 
                    var StudentID = table.Cells[GetAddress(1, studentRow)].Text;
                    if (StudentID == "")
                        break;
                    var stud = new Student() { Id = StudentID, ExcelRow = studentRow };
                    ret.Students.Add(StudentID, stud);
                    if (iColComment != -1)
                        stud.Cohort = table.Cells[GetAddress(iColComment, studentRow)].Text;
                    if (iColComputed != -1)
                        stud.ComputedMark = table.Cells[GetAddress(iColComputed, studentRow)].Text;
                    if (iColMark != -1)
                        stud.Mark = table.Cells[GetAddress(iColMark, studentRow)].Text;

                    // now look at all answers

                    foreach (var question in ret.Questions.Values)
                    {
                        var anstext = table.Cells[GetAddress(question.GetColumn(Question.InformationColumn.Answer), studentRow)].Text;
                        var ansAuto = table.Cells[GetAddress(question.GetColumn(Question.InformationColumn.AutoScore), studentRow)].Text;
                        var ansManual = table.Cells[GetAddress(question.GetColumn(Question.InformationColumn.ManualScore), studentRow)].Text;
                        var answer = new Answer()
                        {
                            QuestionId = question.ID,
                            Text = anstext,
                            AutoScore = MarkFromString(ansAuto),
                            ManualScore = MarkFromString(ansManual),
                        };
                        stud.Answers.Add(question.ID, answer);
                    }
                    studentRow++;
                }
            }
            return ret;
        }

        private static double MarkFromString(string ansAuto)
        {
            if (string.IsNullOrEmpty(ansAuto))
                return -1;
            return Convert.ToDouble(ansAuto);
        }

        private static string GetAddress(int iQcol, int v)
        {
            var first = 0;
            while (iQcol>26)
            {
                iQcol-=26;
                first++;
            }
            if (first > 0)
                return S(first) + S(iQcol) + v;
            return S(iQcol) + v;
        }

        private static string S(int first)
        {
            return ((char)(64 + first)).ToString();
        }

        internal void CleanUp()
        {
            foreach (var q in Questions.Values)
            {
                CleanUp(q);
            }
            foreach (var stud in Students.Values)
            {
                CleanUp(stud);
            }

        }

        private void CleanUp(Student stud)
        {
            foreach (var answer in stud.Answers.Values)
            {
                answer.Text = answer.Text.Replace("</div>,<div>", ",</div><div>").Trim();
                answer.Text = answer.Text.Replace("\\n", "").Trim();
            }
        }

        static private void CleanUp(Question q)
        {
            q.Text = q.Text.Replace("\\n", "").Trim();
        }
    }
}
