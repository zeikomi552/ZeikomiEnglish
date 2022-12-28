using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Common.Converters
{
    /// <summary>
    /// Index値を行番号に変換するコンバーター
    /// </summary>
    [System.Windows.Data.ValueConversion(typeof(int), typeof(int))]
    public class IndexToRowConverter : System.Windows.Data.IValueConverter
    {

        #region IValueConverter メンバ
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var target = (int)value;

            if (target < 0)
            {
                // ここに処理を記述する
                return 0;
            }
            else
            {
                return target + 1;
            }
        }

        // TwoWayの場合に使用する
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        #endregion
    }

}
