using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlanProd.Models
{
    internal class ExcelAdditional : ExcelData
    {
        public string TeacherA { get; set; }
        public string TypeOfWork { get; set; }
        public double? TotalHours { get; set; }
        public ExcelAdditional(string teacher, string typeOfWork, double? totalHours)
        {
            TeacherA=teacher;
            TypeOfWork=typeOfWork;
            TotalHours=totalHours;
        }
        public ExcelAdditional() { }
    }
}
