using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows;
using System.Windows.Controls.Primitives;
using DataGridExtensions;

namespace ProfPlanProd.Views
{
    internal sealed class MultipleChoiceFilter : DataGridExtensions.MultipleChoiceFilter
    {
        static MultipleChoiceFilter()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MultipleChoiceFilter), new FrameworkPropertyMetadata(typeof(DataGridExtensions.MultipleChoiceFilter)));
        }

        public MultipleChoiceFilter()
        {
            SelectAllContent = "Все";
            HasTextFilter = true;
        }
    }
    
    
}
