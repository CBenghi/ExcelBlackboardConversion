using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBlackboardConversion
{
    internal class Reporting
    {
        internal static void ReportByQuestion(Results result, string reportFileName)
        {
            FileInfo f = new FileInfo(reportFileName);
            using var v = f.AppendText();

            HtmlInit(v, "Cohort report");

            foreach (var q in result.Questions.Values)
            {
                v.WriteLine($"<h1>{q.ID} - {q.Text}</h1>");
                v.WriteLine($"<p>Available marks: {q.PossiblePoints}</p>");
                var t = GetProportionalDistribution(result, q).ToList();
                
                var mn = MathNet.Numerics.Statistics.Statistics.Mean(t);
                v.WriteLine($"<p>Mean: <b>{mn*100:0}%</b></p>");
                
                var cov = MathNet.Numerics.Statistics.Statistics.Covariance(t, GetStudentMarkSequence(result, q).ToList());
                v.WriteLine($"<p>Covariance: <b>{cov:0.##}</b></p>");

                v.WriteLine($"<hr />");
                
                v.WriteLine($"<table margin=5 border=1 >");
                v.WriteLine($"<tr>");
                v.WriteLine($"<th>Student ID</th>");
                v.WriteLine($"<th>Answer</th>");
                v.WriteLine($"<th>Score / max</th>");
                v.WriteLine($"<th>This question</th>");
                v.WriteLine($"<th>Entire Module</th>");
                v.WriteLine($"</tr>");
                foreach (var stud in result.Students.Values)
                {
                    if (stud.Answers.TryGetValue(q.ID, out var answer))
                    {
                        v.WriteLine($"<tr>");
                        v.WriteLine($"""<td>{stud.Id}</td>""");
                        v.WriteLine($"""<td>{answer.Text}</td>""");
                        v.WriteLine($"""<td style="text-align: center;">{answer.GetScore()} / {q.PossiblePoints}</td>""");
                        v.WriteLine($"""<td style="text-align: center;">{100*answer.GetScore()/q.PossiblePoints:0}%</td>""");
                        v.WriteLine($"""<td style="text-align: center; color: #aaa;">{stud.Mark}%</td>""");
                        v.WriteLine($"</tr>");
                    }
                }
                v.WriteLine($"</table>");               
            }
            HtmlClose(v);
        }

        private static IEnumerable<double> GetStudentMarkSequence(Results result, Question q)
        {
            foreach (var stud in result.Students.Values)
            {
                yield return Convert.ToDouble(stud.Mark);
                
            }
        }

        private static IEnumerable<double> GetProportionalDistribution(Results result, Question q)
        {
            var available = q.PossiblePoints;
            foreach (var stud in result.Students.Values)
            {
                if (stud.Answers.TryGetValue(q.ID, out var answer))
                {
                    var perc = answer.GetScore();
                    yield return perc / available;
                }
            }
        }

        private static void HtmlClose(StreamWriter v)
        {
            v.WriteLine("</body>");
        }

        private static void HtmlInit(StreamWriter v, string title)
        {
            v.WriteLine("<html>");
            v.WriteLine("<head>");
            v.WriteLine($"<title>{title}</title>");

            v.WriteLine("""
<style>
td, th {
    border: 1px solid black;
    padding: 5px;
}
table {
    border-collapse: collapse;
}
body {
    font-family: Arial, sans-serif;
}
</style>
"""
);


            v.WriteLine("</head>");
            v.WriteLine("<body>");
        }
    }
}
