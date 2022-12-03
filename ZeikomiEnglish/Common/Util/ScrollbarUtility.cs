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
        #region スクロールバー制御
        /// <summary>
        /// スクロールバー制御
        /// 選択している行にスクロールバーを移動する
        /// </summary>
        /// <param name="dg">DataGrid</param>
        public static void TopRow(DataGrid dg)
        {
            dg.ScrollIntoView(dg.Items[dg.SelectedIndex]); // 選択行にスクロールが移動
            dg.UpdateLayout();

        }
        #endregion
    }
}
