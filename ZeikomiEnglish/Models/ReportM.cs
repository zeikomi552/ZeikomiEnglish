using ClosedXML.Excel;
using ControlzEx.Standard;
using Microsoft.Win32;
using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Models
{
    public class ReportM : ModelBase
    {

        #region レポート保存処理
        /// <summary>
        /// レポート保存処理
        /// </summary>
        public static void SaveReport(StoryM story)
        {
            // ダイアログのインスタンスを生成
            var dialog = new SaveFileDialog();

            // ファイルの種類を設定
            dialog.Filter = "エクセルファイル (*.xlsx)|*.xlsx";

            // ダイアログを表示する
            if (dialog.ShowDialog() == true)
            {
                using var wb = new XLWorkbook();
                CreateSummarySheet(wb, "Summary", story);
                CreatePhraseDetailSheet(wb, "Phrase Detail", story);
                CreateWordDetailSheet(wb, "Word Detail", story);
                wb.SaveAs(dialog.FileName);

                ShowMessage.ShowNoticeOK("The report has been output.", "Information");
            }
        }
        #endregion

        #region サマリシートの作成
        /// <summary>
        /// サマリシートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        /// <param name="story">ストーリーオブジェクト</param>
        private static void CreateSummarySheet(XLWorkbook wb, string sheet_name, StoryM story)
        {
            var ws = wb.Worksheets.Add(sheet_name);
            int col = 1;
            ws.Cell(1, col++).Value = "Reigstered Date";                    // 登録日時
            ws.Cell(1, col++).Value = "Total playback time(sec)";           // 合計再生時間
            ws.Cell(1, col++).Value = "Total playback word count";          // 合計再生単語数
            ws.Cell(1, col++).Value = "Total playback phrase count";        // 合計再生フレーズ数
            ws.Cell(1, col++).Value = "Total word translate count";         // 単語検索回数
            ws.Cell(1, col++).Value = "Total phrase translate count";       // フレーズ検索回数
            ws.Cell(1, col++).Value = "Phrase count";                       // フレーズ数
            ws.Cell(1, col++).Value = "Unique phrase translate count";      // フレーズ翻訳回数(ユニーク)
            ws.Cell(1, col++).Value = "Unique word count";                  // 合計単語数(ユニーク)
            ws.Cell(1, col++).Value = "Unique word translate count";        // 合計翻訳単語数(ユニーク)

            col = 1;
            int row = 2;
            ws.Cell(row, col++).Value = DateTime.Now;                         // 現在時刻
            ws.Cell(row, col++).Value = story.TotalElapsedTime;               // 合計再生時間
            ws.Cell(row, col++).Value = story.TotalPlaybackWordCount;         // 合計単語数
            ws.Cell(row, col++).Value = story.TotalPlayBackPhraseCount;       // 合計再生フレーズ数
            ws.Cell(row, col++).Value = story.WordTranslateCount;             // 単語検索回数
            ws.Cell(row, col++).Value = story.PhraseTranslateCount;           // フレーズ検索回数
            ws.Cell(row, col++).Value = story.PhraseItems.Count;              // フレーズ数
            ws.Cell(row, col++).Value = story.GetUniqPhraseTranslateCount();  // フレーズ翻訳回数(ユニーク)
            ws.Cell(row, col++).Value = story.UniqueWordCount;                // 合計単語数(ユニーク)
            ws.Cell(row, col++).Value = story.GetUniqueWordTranslateCount();  // 合計翻訳単語数(ユニーク)

            int col_max = col - 1;  // 最終列を保持

            // スタイルのセット
            SetStyle(ws, col_max, row);
        }
        #endregion

        #region フレーズ詳細シートの作成
        /// <summary>
        /// フレーズ詳細シートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        /// <param name="story">ストーリーオブジェクト</param>
        private static void CreatePhraseDetailSheet(XLWorkbook wb, string sheet_name, StoryM story)
        {
            var ws = wb.Worksheets.Add(sheet_name);
            int col = 1;

            ws.Cell(1, col++).Value = "Word count";                 // 単語数
            ws.Cell(1, col++).Value = "Playback count";             // 再生回数
            ws.Cell(1, col++).Value = "Playback time(sec)";         // 再生時間
            ws.Cell(1, col++).Value = "Phrase translate count";     // 翻訳回数
            ws.Cell(1, col++).Value = "Word translate count";       // 翻訳単語数
            ws.Cell(1, col++).Value = "Phrase";                     // フレーズ

            int col_max = col - 1;  // 最終列を保持

            int row = 2;
            foreach (var tmp in story.PhraseItems.Items)
            {
                col = 1;
                ws.Cell(row, col++).Value = tmp.PlayBackWordCount;  // 単語数
                ws.Cell(row, col++).Value = tmp.PlayCount;          // 再生回数
                ws.Cell(row, col++).Value = tmp.SpeechSec;          // 再生時間
                ws.Cell(row, col++).Value = tmp.TranslateCount;     // 翻訳回数
                ws.Cell(row, col++).Value = (from x in tmp.Words.Items where x.TranslateCount > 0 select x).Count();    // 翻訳単語数
                ws.Cell(row, col++).Value = tmp.Phrase;             // フレーズ
                row++;
            }

            int row_max = row - 1;  // 最終行を保持

            // スタイルのセット
            SetStyle(ws, col_max, row_max);
        }
        #endregion

        #region 単語詳細シートの作成
        /// <summary>
        /// 単語詳細シートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        /// <param name="story">ストーリーオブジェクト</param>
        private static void CreateWordDetailSheet(XLWorkbook wb, string sheet_name, StoryM story)
        {
            Dictionary<string, int> word_list = new Dictionary<string, int>();

            // 単語ごとにまとめる
            foreach (var phrase in story.PhraseItems.Items)
            {
                foreach (var word in phrase.Words.Items)
                {
                    var key = word.Word.ToLower();
                    if (word_list.ContainsKey(key))
                    {
                        word_list[key] += word.TranslateCount;
                    }
                    else
                    {
                        word_list.Add(key, word.TranslateCount);
                    }
                }
            }

            // 単語順にソートする
            var sort_dic = word_list.OrderBy(x => x.Key);

            var ws = wb.Worksheets.Add(sheet_name);
            int col = 1;
            ws.Cell(1, col++).Value = "Word";                 // 単語
            ws.Cell(1, col++).Value = "Translate count";      // 翻訳回数

            int col_max = col - 1;  // 列最大値の保持

            int row = 2;
            foreach (var item in sort_dic)
            {
                ws.Cell(row, 1).Value = item.Key;          // 単語
                ws.Cell(row, 2).Value = item.Value;        // 翻訳回数

                row++;
            }
            int row_max = row - 1; ;

            // スタイルのセット
            SetStyle(ws, col_max, row_max);
        }
        #endregion

        #region スタイル（ヘッダや罫線、幅調整）
        /// <summary>
        /// スタイル（ヘッダや罫線、幅調整）
        /// </summary>
        /// <param name="ws">ワークシート</param>
        /// <param name="col_max">列の最大値</param>
        /// <param name="row_max">行の最大値</param>
        private static void SetStyle(IXLWorksheet ws, int col_max, int row_max)
        {
            ws.Range(ws.Cell(1, 1), ws.Cell(1, col_max)).Style.Fill.SetBackgroundColor(XLColor.FromArgb(0, 0xFF, 0xFF));    // 背景色のセット
            ws.Range(ws.Cell(1, 1), ws.Cell(row_max, col_max)).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);               // 上側の罫線
            ws.Range(ws.Cell(1, 1), ws.Cell(row_max, col_max)).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);            // 下側の罫線
            ws.Range(ws.Cell(1, 1), ws.Cell(row_max, col_max)).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);              // 左側の罫線
            ws.Range(ws.Cell(1, 1), ws.Cell(row_max, col_max)).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);             // 右側の罫線
            ws.Columns(1, col_max).AdjustToContents();                          // 列幅の自動調整
            ws.Range(ws.Cell(1, 1), ws.Cell(1, col_max)).SetAutoFilter();       // オートフィルタのセット
        }
        #endregion
    }
}
