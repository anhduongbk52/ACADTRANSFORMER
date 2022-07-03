using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace ACADTRANSFORMER.Converter
{
    class CheckBoxValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool Checked = (bool)value;
            if (Checked == true)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool check = (bool)value;
            if(check)
            {
                return false;
            }    
            else
            {
                return true;
            } 
        }
    }
}
