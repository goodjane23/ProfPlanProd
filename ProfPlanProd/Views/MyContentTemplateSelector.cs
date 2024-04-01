using ProfPlanProd.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ProfPlanProd.Views
{
    internal class MyContentTemplateSelector : DataTemplateSelector
    {
        public DataTemplate FirstTemplate { get; set; }
        public DataTemplate SecondTemplate { get; set; }
        public DataTemplate ThirdTemplate { get; set; }
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (!(item is TableCollection tableCollection)) return null;

            if (tableCollection.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
            {
                return FirstTemplate;
            }
            else if (tableCollection.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return ThirdTemplate;
            }
            else
            {
                return SecondTemplate;
            }
        }
    }
}
