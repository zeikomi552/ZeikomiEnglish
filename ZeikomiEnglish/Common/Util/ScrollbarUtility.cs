using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ZeikomiEnglish.Common.Util
{
    public class ScrollbarUtility
    {
        public static void TopRow(DataGrid dg)
        {
            dg.ScrollIntoView(dg.Items[dg.SelectedIndex]); // 選択行にスクロールが移動
            dg.UpdateLayout();

        }
    }
}
