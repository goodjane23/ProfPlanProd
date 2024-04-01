using ProfPlanProd.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlanProd.Models
{
    internal class ExcelData : ViewModel
    {
        public string GetTermValue()
        {
            // Проверяем, является ли текущий объект экземпляром класса ExcelModel
            if (this is ExcelModel excelModel)
            {
                // Если объект является экземпляром ExcelModel, возвращаем значение свойства Term
                return excelModel.Term;
            }
            else
            {
                // Если объект не является экземпляром ExcelModel, возвращаем null или другое значение по умолчанию
                return "unnull"; // или можно вернуть другое значение, если требуется
            }
        }
    }
}
