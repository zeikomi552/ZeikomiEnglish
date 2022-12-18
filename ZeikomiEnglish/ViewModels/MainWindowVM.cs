using ClosedXML.Excel;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Win32;
using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using MVVMCore.Common.Wrapper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ZeikomiEnglish.Common.Util;
using ZeikomiEnglish.Models;

namespace ZeikomiEnglish.ViewModels
{
    public class MainWindowVM : ViewModelBase
    {

        #region 単一フレーズの繰り返し[IsPressSinglePhrase]プロパティ
        /// <summary>
        /// 単一フレーズの繰り返し[IsPressSinglePhrase]プロパティ用変数
        /// </summary>
        bool _IsPressSinglePhrase = false;
        /// <summary>
        /// 単一フレーズの繰り返し[IsPressSinglePhrase]プロパティ
        /// </summary>
        public bool IsPressSinglePhrase
        {
            get
            {
                return _IsPressSinglePhrase;
            }
            set
            {
                if (!_IsPressSinglePhrase.Equals(value))
                {
                    _IsPressSinglePhrase = value;
                    NotifyPropertyChanged("IsPressSinglePhrase");
                }
            }
        }
        #endregion

        #region 音声再生命令中[IsPressVoice]プロパティ
        /// <summary>
        /// 音声再生命令中[IsPressVoice]プロパティ用変数
        /// </summary>
        bool _IsPressVoice = false;
        /// <summary>
        /// 音声再生命令中[IsPressVoice]プロパティ
        /// </summary>
        public bool IsPressVoice
        {
            get
            {
                return _IsPressVoice;
            }
            set
            {
                if (!_IsPressVoice.Equals(value))
                {
                    _IsPressVoice = value;
                    NotifyPropertyChanged("IsPressVoice");
                }
            }
        }
        #endregion

        #region スピーチのレート[SpeechRate]プロパティ
        /// <summary>
        /// スピーチのレート[SpeechRate]プロパティ用変数
        /// </summary>
        int _SpeechRate = 0;
        /// <summary>
        /// スピーチのレート[SpeechRate]プロパティ
        /// </summary>
        public int SpeechRate
        {
            get
            {
                return _SpeechRate;
            }
            set
            {
                if (!_SpeechRate.Equals(value))
                {
                    _SpeechRate = value;
                    NotifyPropertyChanged("SpeechRate");
                }
            }
        }
        #endregion

        #region フレーズ要素[PhraseItems]プロパティ
        /// <summary>
        /// フレーズ要素[PhraseItems]プロパティ用変数
        /// </summary>
        ModelList<PhraseM> _PhraseItems = new ModelList<PhraseM>();
        /// <summary>
        /// フレーズ要素[PhraseItems]プロパティ
        /// </summary>
        public ModelList<PhraseM> PhraseItems
        {
            get
            {
                return _PhraseItems;
            }
            set
            {
                if (_PhraseItems == null || !_PhraseItems.Equals(value))
                {
                    _PhraseItems = value;
                    NotifyPropertyChanged("PhraseItems");
                }
            }
        }
        #endregion

        #region テキスト[Text]プロパティ
        /// <summary>
        /// テキスト[Text]プロパティ用変数
        /// </summary>
        string _Text = string.Empty;
        /// <summary>
        /// テキスト[Text]プロパティ
        /// </summary>
        public string Text
        {
            get
            {
                return _Text;
            }
            set
            {
                if (_Text == null || !_Text.Equals(value))
                {
                    _Text = value;
                    NotifyPropertyChanged("Text");
                }
            }
        }
        #endregion

        #region true:英英辞典 false:Google翻訳[EngDictionaryF]プロパティ
        /// <summary>
        /// true:英英辞典 false:Google翻訳[EngDictionaryF]プロパティ用変数
        /// </summary>
        bool _EngDictionaryF = false;
        /// <summary>
        /// true:英英辞典 false:Google翻訳[EngDictionaryF]プロパティ
        /// </summary>
        public bool EngDictionaryF
        {
            get
            {
                return _EngDictionaryF;
            }
            set
            {
                if (!_EngDictionaryF.Equals(value))
                {
                    _EngDictionaryF = value;
                    NotifyPropertyChanged("EngDictionaryF");
                }
            }
        }
        #endregion

        #region true:DeepL false:Google翻訳[DeepL_F]プロパティ
        /// <summary>
        /// true:DeepL false:Google翻訳[DeepL_F]プロパティ用変数
        /// </summary>
        bool _DeepL_F = true;
        /// <summary>
        /// true:DeepL false:Google翻訳[DeepL_F]プロパティ
        /// </summary>
        public bool DeepL_F
        {
            get
            {
                return _DeepL_F;
            }
            set
            {
                if (!_DeepL_F.Equals(value))
                {
                    _DeepL_F = value;
                    NotifyPropertyChanged("DeepL_F");
                }
            }
        }
        #endregion


        #region インストールされている音声リスト[VoiceList]プロパティ
        /// <summary>
        /// インストールされている音声リスト[VoiceList]プロパティ用変数
        /// </summary>
        ModelList<InstalledVoice> _VoiceList = new ModelList<InstalledVoice>();
        /// <summary>
        /// インストールされている音声リスト[VoiceList]プロパティ
        /// </summary>
        public ModelList<InstalledVoice> VoiceList
        {
            get
            {
                return _VoiceList;
            }
            set
            {
                if (_VoiceList == null || !_VoiceList.Equals(value))
                {
                    _VoiceList = value;
                    NotifyPropertyChanged("VoiceList");
                }
            }
        }
        #endregion
        #region 経過時間（再生時間のみ）[TotalElapsedTime]プロパティ
        /// <summary>
        /// 経過時間（再生時間のみ）[TotalElapsedTime]プロパティ用変数
        /// </summary>
        double _TotalElapsedTime = 0.0;
        /// <summary>
        /// 経過時間（再生時間のみ）[TotalElapsedTime]プロパティ
        /// </summary>
        public double TotalElapsedTime
        {
            get
            {
                return _TotalElapsedTime;
            }
            set
            {
                if (!_TotalElapsedTime.Equals(value))
                {
                    _TotalElapsedTime = value;
                    NotifyPropertyChanged("TotalElapsedTime");
                }
            }
        }
        #endregion

        #region 合計再生単語数[TotalWordCount]プロパティ
        /// <summary>
        /// 合計再生単語数[TotalWordCount]プロパティ用変数
        /// </summary>
        int _TotalWordCount = 0;
        /// <summary>
        /// 合計再生単語数[TotalWordCount]プロパティ
        /// </summary>
        public int TotalWordCount
        {
            get
            {
                return _TotalWordCount;
            }
            set
            {
                if (!_TotalWordCount.Equals(value))
                {
                    _TotalWordCount = value;
                    NotifyPropertyChanged("TotalWordCount");
                }
            }
        }
        #endregion



        #region キャッシュの保存先ディレクトリ
        /// <summary>
        /// キャッシュの保存先ディレクトリ
        /// </summary>
        private string _WebViewDir = "EBWebView";
        #endregion

        #region 初期化処理(WebView2の配布)
        /// <summary>
        /// 初期化処理(WebView2の配布)
        /// </summary>
        private async void InitializeAsync(MainWindow wnd)
        {
            var browserExecutableFolder = Path.Combine(MVVMCore.Common.Utilities.PathManager.GetApplicationFolder(), _WebViewDir);

            // カレントディレクトリの作成
            MVVMCore.Common.Utilities.PathManager.CreateDirectory(browserExecutableFolder);

            // 環境の作成
            var webView2Environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, browserExecutableFolder);

            // 固定バージョンのブラウザを配布
            await wnd.WebView2Ctrl.EnsureCoreWebView2Async(webView2Environment);
        }
        #endregion

        #region 初期化処理
        /// <summary>
        /// 初期化処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public override void Init(object sender, EventArgs e)
        {
            try
            {
                var wnd = sender as MainWindow;
                if (wnd != null)
                {
                    try
                    {
                        InitializeAsync(wnd);
                    }
                    catch
                    {
                        ShowMessage.ShowNoticeOK("WebView2ランタイムがインストールされていないようです。\r\nインストールしてください", "通知");
                        URLUtility.OpenUrl("https://developer.microsoft.com/en-us/microsoft-edge/webview2/");
                    }
                }


                this.VoiceList.Items.Clear();
                var synthesizer = new SpeechSynthesizer();
                var installedVoices = synthesizer.GetInstalledVoices();

                foreach (var voice in installedVoices)
                {
                    this.VoiceList.Items.Add(voice);
                }

                // nullチェック
                if (this.VoiceList.Items.FirstOrDefault() != null)
                {
                    var tmp = (from x in this.VoiceList.Items
                               where x.VoiceInfo.Name.Contains("Zira")
                               select x).FirstOrDefault();

                    if (tmp == null)
                    {
                        this.VoiceList.SelectedItem = this.VoiceList.Items.FirstOrDefault()!;
                    }
                    else
                    {
                        this.VoiceList.SelectedItem = tmp;
                    }
                }

            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
        }
        #endregion

        #region 閉じる処理
        /// <summary>
        /// 閉じる処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public override void Close(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
        }
        #endregion

        #region フレーズリストの選択が変化した場合の処理
        /// <summary>
        /// フレーズリストの選択が変化した場合の処理
        /// </summary>
        public void TextChanged()
        {
            // フレーズに分解する
            var list = this.Text.Replace("\r", "").Split("\n");

            // フレーズリストをクリア
            this.PhraseItems.Items.Clear();

            // リスト数まわす
            foreach (var tmp in list)
            {
                if (tmp.Trim().Length > 0)
                {
                    // フレーズリストに追加
                    this.PhraseItems.Items.Add(new PhraseM()
                    {
                        Phrase = tmp.Trim()
                    }
                    );
                }
            }
        }
        #endregion

        #region フレーズのダブルクリック
        /// <summary>
        /// フレーズのダブルクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void OpenURL(object sender, EventArgs e)
        {
            try
            {
                // ウィンドウを取得
                var wnd = VisualTreeHelperWrapper.GetWindow<MainWindow>(sender) as MainWindow;

                // nullチェック
                if (wnd != null)
                {
                    var url_base = "https://translate.google.co.jp/?sl=en&tl=ja&text={0}&op=translate";   // Google翻訳のURL
                    if (this.DeepL_F)
                    {
                        url_base = "https://www.deepl.com/ja/translator#en/ja/{0}";   // DeepLのURL
                    }


                    // nullチェック
                    if (this.PhraseItems.SelectedItem != null)
                    {
                        string url = string.Format(url_base, this.PhraseItems.SelectedItem.Phrase);   // URL作成

                        // nullチェック
                        if (wnd.WebView2Ctrl != null && wnd.WebView2Ctrl.CoreWebView2 != null)
                        {
                            // URLを開く
                            wnd.WebView2Ctrl.CoreWebView2.Navigate(url);
                        }
                    }
                }
            }
            catch
            {
            }
        }
        #endregion

        #region 単語のダブルクリック
        /// <summary>
        /// 単語のダブルクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void WordDoubleClick(object sender, EventArgs e)
        {
            try
            {
                // ウィンドウを取得
                var wnd = VisualTreeHelperWrapper.GetWindow<MainWindow>(sender) as MainWindow;

                // nullチェック
                if (wnd != null)
                {
                    var url_base = "https://translate.google.co.jp/?sl=en&tl=ja&text={0}&op=translate";

                    if (this.EngDictionaryF)
                    {
                        url_base = "https://dictionary.cambridge.org/dictionary/english/{0}";
                    }

                    // nullチェック
                    if (this.PhraseItems.SelectedItem != null && this.PhraseItems.SelectedItem.Words.SelectedItem != null)
                    {
                        string url = string.Format(url_base, this.PhraseItems.SelectedItem.Words.SelectedItem.Word);   // URL作成

                        // nullチェック
                        if (wnd.WebView2Ctrl != null && wnd.WebView2Ctrl.CoreWebView2 != null)
                        {
                            // URLを開く
                            wnd.WebView2Ctrl.CoreWebView2.Navigate(url);
                        }
                    }
                }
            }
            catch
            {
            }
        }
        #endregion

        #region 行分割
        /// <summary>
        /// 行分割
        /// </summary>
        public void PeriodLineBreak()
        {
            try
            {
                if (this.Text != null)
                {
                    this.Text = this.Text.Replace("\r", "")
                        .Replace("\n", "")
                        .Replace(".", ".\r\n")
                        .Replace("!", "!\r\n")
                        .Replace("?", "?\r\n")
                        .Replace(".\r\n\"", ".\"\r\n")
                        .Replace("!\r\n\"", "!\"\r\n")
                        .Replace("?\r\n\"", "?\"\r\n")
                        .Replace(";",";\r\n");
                }
            }
            catch
            {

            }
        }
        #endregion

        #region 選択行が変化した際の処理
        /// <summary>
        /// 選択行が変化した際の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void SelectedItemChanged(object sender, EventArgs e)
        {
            try
            {
                // ウィンドウを取得
                var wnd = VisualTreeHelperWrapper.GetWindow<MainWindow>(sender) as MainWindow;

                // ウィンドウが取得できた場合
                if (wnd != null)
                {
                    ScrollbarUtility.TopRow(wnd.phrase_dg);
                }
            }
            catch
            {

            }
        }
        #endregion

        #region フレーズを音声再生する
        /// <summary>
        /// フレーズを音声再生する
        /// </summary>
        public void PhraseVoiceSingle()
        {
            try
            {
                Task.Run(() =>
                {
                    while (this.IsPressSinglePhrase)
                    {
                        if (this.PhraseItems.Count <= 0)
                        {
                            this.IsPressSinglePhrase = false;
                            return;
                        }

                        // nullチェック
                        if (this.PhraseItems.SelectedItem == null)
                        {
                            this.PhraseItems.SelectedItem = this.PhraseItems.First();
                        }

                        // 音声再生
                        int index = this.PhraseItems.IndexOf(this.PhraseItems.SelectedItem);
                        var tm = VoiceUtility.PhraseVoice(this.PhraseItems.SelectedItem!.Phrase, this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate);  // フレーズ再生
                        
                        // オブジェクトに保存
                        var phrase_tmp = this.PhraseItems.ElementAt(index);

                        if (tm.TotalSeconds > 0)
                        {
                            this.TotalElapsedTime += phrase_tmp.SpeechSec = tm.TotalSeconds;    // 再生時間保存
                            this.TotalWordCount += phrase_tmp.WordCount;                        // 合計再生単語数
                            phrase_tmp.PlayCount++;                 // 再生回数インクリメント
                        }
                        else
                        {
                            // 再生時間が0(スリープに入ってしまった可能性がある)ため抜ける
                            this.IsPressSinglePhrase = false;
                        }
                    }
                });
            }
            catch
            {
            }
        }
        #endregion

        #region 合成音声を連続で再生する
        /// <summary>
        /// 合成音声を連続で再生する
        /// </summary>
        public void PhraseVoiceMulti()
        {
            try
            {
                Task.Run(() =>
                {
                    // nullチェック
                    if (this.PhraseItems.SelectedItem == null)
                    {
                        this.PhraseItems.SelectedItem = this.PhraseItems.First();
                    }

                    // 音声再生インデックス取得
                    int index = this.PhraseItems.IndexOf(this.PhraseItems.SelectedItem);

                    while (this.IsPressVoice)
                    {
                        if (this.PhraseItems.Count <= 0)
                        {
                            this.IsPressVoice = false;
                            return;
                        }

                        if (index < this.PhraseItems.Count)
                        {
                            var tmp = this.PhraseItems.ElementAt(index);

                            this.PhraseItems.SelectedItem = tmp;
                            System.Threading.Thread.Sleep(100);
                            var tm = VoiceUtility.PhraseVoice(tmp.Phrase, this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate);    // フレーズ再生

                            if (this.PhraseItems.Count > index && this.PhraseItems.ElementAt(index) != null)
                            {
                                // オブジェクトに保存
                                var phrase_tmp = this.PhraseItems.ElementAt(index);
                                if(tm.TotalSeconds > 0)
                                {
                                    this.TotalElapsedTime += phrase_tmp.SpeechSec = tm.TotalSeconds;    // 再生時間保存
                                    this.TotalWordCount += phrase_tmp.WordCount;                        // 合計再生単語数
                                    phrase_tmp.PlayCount++;                 // 再生回数インクリメント
                                }
                                else
                                {
                                    // 再生時間が0(スリープに入ってしまった可能性がある)ため抜ける
                                    this.IsPressVoice = false;
                                }
                            }
                            index++;
                        }
                        else
                        {
                            index = 0;
                        }
                    }
                });
            }
            catch
            {
            }
        }
        #endregion

        #region 音声の録音
        /// <summary>
        /// 合成音声の録音
        /// </summary>
        public void RecordVoice()
        {
            try
            {
                if (this.PhraseItems.Count <= 0)
                {
                    this.IsPressSinglePhrase = false;
                    return;
                }

                // ダイアログのインスタンスを生成
                var dialog = new SaveFileDialog();

                // ファイルの種類を設定
                dialog.Filter = "テキストファイル (*.wav)|*.wav";

                // ダイアログを表示する
                if (dialog.ShowDialog() == true)
                {
                    // nullチェック
                    if (this.PhraseItems.SelectedItem == null)
                    {
                        this.PhraseItems.SelectedItem = this.PhraseItems.First();
                    }

                    // 音声再生
                    int index = this.PhraseItems.IndexOf(this.PhraseItems.SelectedItem);
                    var tm = VoiceUtility.PhraseVoiceToFile(this.PhraseItems.SelectedItem!.Phrase, this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate, dialog.FileName);  // フレーズ再生

                    // オブジェクトに保存
                    var phrase_tmp = this.PhraseItems.ElementAt(index);

                    if (tm.TotalSeconds > 0)
                    {
                        this.TotalElapsedTime += phrase_tmp.SpeechSec = tm.TotalSeconds;    // 再生時間保存
                        this.TotalWordCount += phrase_tmp.WordCount;                        // 合計再生単語数
                        phrase_tmp.PlayCount++;                 // 再生回数インクリメント
                    }
                    else
                    {
                        // 再生時間が0(スリープに入ってしまった可能性がある)ため抜ける
                        this.IsPressSinglePhrase = false;
                    }
                }                
            }
            catch
            {
            }
        }
        #endregion

        /// <summary>
        /// 合成音声の録音
        /// </summary>
        public void RecordVoiceMulti()
        {
            try
            {
                StringBuilder text = new StringBuilder();
                foreach(var tmp in this.PhraseItems.Items)
                {
                    text.AppendLine(tmp.Phrase);
                }

                // ダイアログのインスタンスを生成
                var dialog = new SaveFileDialog();

                // ファイルの種類を設定
                dialog.Filter = "テキストファイル (*.wav)|*.wav";

                // ダイアログを表示する
                if (dialog.ShowDialog() == true)
                {
                    VoiceUtility.PhraseVoiceToFile(text.ToString(), this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate, dialog.FileName);  // フレーズ再生
                }
            }
            catch
            {
            }
        }

        #region レポート保存処理
        /// <summary>
        /// レポート保存処理
        /// </summary>
        public void SaveReport()
        {
            try
            {
                // ダイアログのインスタンスを生成
                var dialog = new SaveFileDialog();

                // ファイルの種類を設定
                dialog.Filter = "エクセルファイル (*.xlsx)|*.xlsx";

                // ダイアログを表示する
                if (dialog.ShowDialog() == true)
                {
                    using var wb = new XLWorkbook();
                    CreateSummarySheet(wb, "Summary");
                    CreateDetailSheet(wb, "Detail");
                    wb.SaveAs(dialog.FileName);

                    ShowMessage.ShowNoticeOK("The report has been output.", "Information");
                }

            }
            catch
            {
            }
        }
        #endregion

        #region サマリシートの作成
        /// <summary>
        /// サマリシートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        private void CreateSummarySheet(XLWorkbook wb, string sheet_name)
        {
            var ws = wb.Worksheets.Add(sheet_name);
            ws.Cell(1, 1).Value = "Reigstered Date";            // 登録日時
            ws.Cell(1, 2).Value = "Total playback time(sec)";   // 合計再生時間
            ws.Cell(1, 3).Value = "Total word count";           // 合計単語数

            ws.Cell(2, 1).Value = DateTime.Now;                 // 現在時刻
            ws.Cell(2, 2).Value = this.TotalElapsedTime;        // 合計再生時間
            ws.Cell(2, 3).Value = this.TotalWordCount;          // 合計単語数
        }
        #endregion

        #region 詳細シートの作成
        /// <summary>
        /// 詳細シートの作成
        /// </summary>
        /// <param name="wb">ワークブックオブジェクト</param>
        /// <param name="sheet_name">シート名</param>
        private void CreateDetailSheet(XLWorkbook wb, string sheet_name)
        {
            var ws = wb.Worksheets.Add(sheet_name);
            ws.Cell(1, 1).Value = "Word count";                 // 単語数
            ws.Cell(1, 2).Value = "Playback count";             // 再生回数
            ws.Cell(1, 3).Value = "Playback time(sec)";         // 再生時間
            ws.Cell(1, 4).Value = "Phrase";                     // フレーズ

            int row = 2;
            foreach(var tmp in this.PhraseItems.Items)
            {
                ws.Cell(row, 1).Value = tmp.WordCount;          // 単語数
                ws.Cell(row, 2).Value = tmp.PlayCount;          // 再生回数
                ws.Cell(row, 3).Value = tmp.SpeechSec;          // 再生時間
                ws.Cell(row, 4).Value = tmp.Phrase;             // フレーズ

                row++;
            }
        }
        #endregion
    }
}
