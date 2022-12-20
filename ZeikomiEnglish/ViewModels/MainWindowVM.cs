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
        #region ストーリーオブジェクト[Story]プロパティ
        /// <summary>
        /// ストーリーオブジェクト[Story]プロパティ用変数
        /// </summary>
        StoryM _Story = new StoryM();
        /// <summary>
        /// ストーリーオブジェクト[Story]プロパティ
        /// </summary>
        public StoryM Story
        {
            get
            {
                return _Story;
            }
            set
            {
                if (_Story == null || !_Story.Equals(value))
                {
                    _Story = value;
                    NotifyPropertyChanged("Story");
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


                this.Story.VoiceList.Items.Clear();
                var synthesizer = new SpeechSynthesizer();
                var installedVoices = synthesizer.GetInstalledVoices();

                foreach (var voice in installedVoices)
                {
                    this.Story.VoiceList.Items.Add(voice);
                }

                // nullチェック
                if (this.Story.VoiceList.Items.FirstOrDefault() != null)
                {
                    var tmp = (from x in this.Story.VoiceList.Items
                               where x.VoiceInfo.Name.Contains("Zira")
                               select x).FirstOrDefault();

                    if (tmp == null)
                    {
                        this.Story.VoiceList.SelectedItem = this.Story.VoiceList.Items.FirstOrDefault()!;
                    }
                    else
                    {
                        this.Story.VoiceList.SelectedItem = tmp;
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
            var list = this.Story.Text.Replace("\r", "").Split("\n");

            // フレーズリストをクリア
            this.Story.PhraseItems.Items.Clear();

            // リスト数まわす
            foreach (var tmp in list)
            {
                if (tmp.Trim().Length > 0)
                {
                    // フレーズリストに追加
                    this.Story.PhraseItems.Items.Add(new PhraseM()
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
                    if (this.Story.PhraseItems.SelectedItem != null)
                    {
                        string url = string.Format(url_base, this.Story.PhraseItems.SelectedItem.Phrase);   // URL作成

                        // nullチェック
                        if (wnd.WebView2Ctrl != null && wnd.WebView2Ctrl.CoreWebView2 != null)
                        {
                            // URLを開く
                            wnd.WebView2Ctrl.CoreWebView2.Navigate(url);

                            // フレーズ検索回数
                            this.Story.PhraseSearch++;
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
                    if (this.Story.PhraseItems.SelectedItem != null && this.Story.PhraseItems.SelectedItem.Words.SelectedItem != null)
                    {
                        string url = string.Format(url_base, this.Story.PhraseItems.SelectedItem.Words.SelectedItem.Word);   // URL作成

                        // nullチェック
                        if (wnd.WebView2Ctrl != null && wnd.WebView2Ctrl.CoreWebView2 != null)
                        {
                            // URLを開く
                            wnd.WebView2Ctrl.CoreWebView2.Navigate(url);

                            // 単語検索回数インクリメント
                            this.Story.WordSearch++;
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
                if (this.Story.Text != null)
                {
                    this.Story.Text = this.Story.Text.Replace("\r", "")
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
                    while (this.Story.IsPressSinglePhrase)
                    {
                        if (this.Story.PhraseItems.Count <= 0)
                        {
                            this.Story.IsPressSinglePhrase = false;
                            return;
                        }

                        // nullチェック
                        if (this.Story.PhraseItems.SelectedItem == null)
                        {
                            this.Story.PhraseItems.SelectedItem = this.Story.PhraseItems.First();
                        }

                        // 音声再生
                        int index = this.Story.PhraseItems.IndexOf(this.Story.PhraseItems.SelectedItem);
                        var tm = RecordM.PhraseVoice(this.Story.PhraseItems.SelectedItem!.Phrase, this.Story.VoiceList.SelectedItem.VoiceInfo.Name, this.Story.SpeechRate);  // フレーズ再生
                        
                        // オブジェクトに保存
                        var phrase_tmp = this.Story.PhraseItems.ElementAt(index);

                        if (tm.TotalSeconds > 0)
                        {
                            this.Story.TotalElapsedTime += phrase_tmp.SpeechSec = tm.TotalSeconds;    // 再生時間保存
                            this.Story.TotalWordCount += phrase_tmp.WordCount;                        // 合計再生単語数
                            phrase_tmp.PlayCount++;                 // 再生回数インクリメント
                        }
                        else
                        {
                            // 再生時間が0(スリープに入ってしまった可能性がある)ため抜ける
                            this.Story.IsPressSinglePhrase = false;
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
                    if (this.Story.PhraseItems.SelectedItem == null)
                    {
                        this.Story.PhraseItems.SelectedItem = this.Story.PhraseItems.First();
                    }

                    // 音声再生インデックス取得
                    int index = this.Story.PhraseItems.IndexOf(this.Story.PhraseItems.SelectedItem);

                    while (this.Story.IsPressVoice)
                    {
                        if (this.Story.PhraseItems.Count <= 0)
                        {
                            this.Story.IsPressVoice = false;
                            return;
                        }

                        if (index < this.Story.PhraseItems.Count)
                        {
                            var tmp = this.Story.PhraseItems.ElementAt(index);

                            this.Story.PhraseItems.SelectedItem = tmp;
                            System.Threading.Thread.Sleep(100);
                            var tm = RecordM.PhraseVoice(tmp.Phrase, this.Story.VoiceList.SelectedItem.VoiceInfo.Name, this.Story.SpeechRate);    // フレーズ再生

                            if (this.Story.PhraseItems.Count > index && this.Story.PhraseItems.ElementAt(index) != null)
                            {
                                // オブジェクトに保存
                                var phrase_tmp = this.Story.PhraseItems.ElementAt(index);
                                if(tm.TotalSeconds > 0)
                                {
                                    this.Story.TotalElapsedTime += phrase_tmp.SpeechSec = tm.TotalSeconds;    // 再生時間保存
                                    this.Story.TotalWordCount += phrase_tmp.WordCount;                        // 合計再生単語数
                                    phrase_tmp.PlayCount++;                 // 再生回数インクリメント
                                }
                                else
                                {
                                    // 再生時間が0(スリープに入ってしまった可能性がある)ため抜ける
                                    this.Story.IsPressVoice = false;
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
                if (this.Story.PhraseItems.Count <= 0)
                {
                    this.Story.IsPressSinglePhrase = false;
                    return;
                }

                // nullチェック
                if (this.Story.PhraseItems.SelectedItem == null)
                {
                    this.Story.PhraseItems.SelectedItem = this.Story.PhraseItems.First();
                }

                // 1行の録音処理
                RecordM.RecordSingle(this.Story, this.Story.VoiceList.SelectedItem.VoiceInfo.Name, this.Story.SpeechRate);
            }
            catch
            {
            }
        }
        #endregion

        #region 合成音声の録音
        /// <summary>
        /// 合成音声の録音
        /// </summary>
        public void RecordVoiceMulti()
        {
            try
            {
                // 複数行の合成音声の録音処理
                RecordM.RecordVoiceMulti(this.Story, this.Story.VoiceList.SelectedItem.VoiceInfo.Name, this.Story.SpeechRate);
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
        }
        #endregion

        #region レポートの保存処理
        /// <summary>
        /// レポートの保存処理
        /// </summary>
        public void SaveReport()
        {
            try
            {
                ReportM.SaveReport(this.Story);
            }
            catch(Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
        }
        #endregion
    }
}
