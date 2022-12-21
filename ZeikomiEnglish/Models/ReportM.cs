using ClosedXML.Excel;
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
            ws.Cell(1, 1).Value = "Reigstered Date";                    // 登録日時
            ws.Cell(1, 2).Value = "Total playback time(sec)";           // 合計再生時間
            ws.Cell(1, 3).Value = "Total word count";                   // 合計単語数
            ws.Cell(1, 4).Value = "Total word translate count";         // 単語検索回数
            ws.Cell(1, 5).Value = "Total phrase translate count";       // フレーズ検索回数

            ws.Cell(2, 1).Value = DateTime.Now;                         // 現在時刻
            ws.Cell(2, 2).Value = story.TotalElapsedTime;               // 合計再生時間
            ws.Cell(2, 3).Value = story.TotalWordCount;                 // 合計単語数
            ws.Cell(2, 4).Value = story.WordTranslateCount;             // 単語検索回数
            ws.Cell(2, 5).Value = story.PhraseTranslateCount;           // フレーズ検索回数
        }
        #endregion

        #region 詳細シートの作成
        /// <summary>
        /// 詳細シートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        /// <param name="story">ストーリーオブジェクト</param>
        private static void CreatePhraseDetailSheet(XLWorkbook wb, string sheet_name, StoryM story)
        {
            var ws = wb.Worksheets.Add(sheet_name);
            ws.Cell(1, 1).Value = "Word count";                 // 単語数
            ws.Cell(1, 2).Value = "Playback count";             // 再生回数
            ws.Cell(1, 3).Value = "Playback time(sec)";         // 再生時間
            ws.Cell(1, 4).Value = "Phrase";                     // フレーズ
            ws.Cell(1, 5).Value = "Phrase translate count";     // 翻訳回数

            int row = 2;
            foreach (var tmp in story.PhraseItems.Items)
            {
                ws.Cell(row, 1).Value = tmp.WordCount;          // 単語数
                ws.Cell(row, 2).Value = tmp.PlayCount;          // 再生回数
                ws.Cell(row, 3).Value = tmp.SpeechSec;          // 再生時間
                ws.Cell(row, 4).Value = tmp.Phrase;             // フレーズ
                ws.Cell(row, 5).Value = tmp.TranslateCount;     // 翻訳回数

                row++;
            }
        }
        #endregion


        /// <summary>
        /// 詳細シートの作成
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
            ws.Cell(1, 1).Value = "Word";                 // 単語
            ws.Cell(1, 2).Value = "Word translate count";      // 翻訳回数

            int row = 2;
            foreach (var item in sort_dic)
            {
                ws.Cell(row, 1).Value = item.Key;          // 単語
                ws.Cell(row, 2).Value = item.Value;        // 翻訳回数

                row++;
            }
        }

    }
}
