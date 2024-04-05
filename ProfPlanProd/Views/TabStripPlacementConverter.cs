using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace ProfPlanProd.Views
{
    internal class TabStripPlacementConverter : IValueConverter
    {
        public ControlTemplate TopTemplate { get; set; }
        public ControlTemplate BottomTemplate { get; set; }
        public ControlTemplate LeftTemplate { get; set; }
        public ControlTemplate RightTemplate { get; set; }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var dockValue = (Dock)value;
            
            switch (dockValue)
            {
                case Dock.Left:
                    return LeftTemplate;                   
                case Dock.Top:
                    return TopTemplate;                    
                case Dock.Right:
                    return RightTemplate;
                case Dock.Bottom:
                    return BottomTemplate;
                default:
                    return RightTemplate;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    
    }
}
