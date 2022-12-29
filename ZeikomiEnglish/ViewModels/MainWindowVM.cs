using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Win32;
using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using MVVMCore.Common.Wrapper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        StoryM _Story = new ();
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
            try
            {
                var browserExecutableFolder = Path.Combine(MVVMCore.Common.Utilities.PathManager.GetApplicationFolder(), _WebViewDir);

                // カレントディレクトリの作成
                MVVMCore.Common.Utilities.PathManager.CreateDirectory(browserExecutableFolder);

                // 環境の作成
                var webView2Environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, browserExecutableFolder);

                // 固定バージョンのブラウザを配布
                await wnd.WebView2Ctrl.EnsureCoreWebView2Async(webView2Environment);
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
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

                // 合成音声の初期化
                this.Story.InitVoice();
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
                // アプリケーションの終了確認
                if (ShowMessage.ShowQuestionYesNo("Do you want to close the application?", "Question") == MessageBoxResult.No)
                {
                    if (e != null && (e as CancelEventArgs) != null)
                    {
                        ((CancelEventArgs)e).Cancel = true; // 終了キャンセル
                    }
                }
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
            try
            {
                // フレーズリストの更新
                this.Story.RefreshPhraseList();
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                    this.Story.TranslatePhrase(wnd.WebView2Ctrl);
                }
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                    this.Story.TranslateWord(wnd.WebView2Ctrl);
                }
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                // 文章の整形
                this.Story.PeriodLineBreak();
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                if (wnd != null && this.Story.PhraseItems.SelectedItem != null)
                {
                    ScrollbarUtility.TopRow(wnd.phrase_dg);
                }
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                this.Story.PhraseVoiceSingle();
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
                this.Story.PhraseVoiceMulti();
            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
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

        #region キー入力処理の受付
        /// <summary>
        /// キー入力処理の受付
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void KeyDown(object sender, EventArgs e)
        {
            try
            {
                var wnd = sender as MainWindow;

                if (wnd != null)
                {
                    var key_eve = e as KeyEventArgs;

                    if (key_eve!.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || key_eve!.KeyboardDevice.IsKeyDown(Key.RightCtrl))
                    {
                        switch (key_eve.Key)
                        {
                            case Key.Left:
                                {
                                    if (this.Story.PhraseItems.Count > 0)
                                    {
                                        var targetGrid = wnd.phrase_dg;

                                        // 未選択状態なので1行目を選択する
                                        if (this.Story.PhraseItems.SelectedItem == null)
                                        {
                                            targetGrid.SelectedIndex = 0;
                                        }

                                        // 最初のセルを取得
                                        var cellInfo = targetGrid.SelectedCells.FirstOrDefault();

                                        // フォーカスのセット
                                        targetGrid.Focus();

                                        // カレントセルのセット
                                        targetGrid.CurrentCell = new DataGridCellInfo();
                                        targetGrid.CurrentCell = cellInfo;
                                    }
                                    break;
                                }
                            case Key.Right:
                                {
                                    if (this.Story.PhraseItems.SelectedItem != null
                                        && this.Story.PhraseItems.SelectedItem.Words.Items.Count > 0)
                                    {
                                        var targetGrid = wnd.word_dg;

                                        // 行が選択されていない場合最初の行を選択
                                        if (this.Story.PhraseItems.SelectedItem.Words.SelectedItem == null)
                                        {
                                            targetGrid.SelectedIndex = 0;
                                        }

                                        // 最初のセルを取得
                                        var cellInfo = targetGrid.SelectedCells.FirstOrDefault();

                                        // フォーカスをセット
                                        targetGrid.Focus();

                                        // カレントセルのセット
                                        targetGrid.CurrentCell = new DataGridCellInfo();
                                        targetGrid.CurrentCell = cellInfo;
                                    }
                                    break;
                                }

                        }
                    }

                    switch (key_eve.Key)
                    {
                        case Key.Enter:
                            {
                                var active = FocusManager.GetFocusedElement(wnd);

                                if (active is DataGridCell)
                                {
                                    var tmp = ((DataGridCell)active).DataContext as WordM;

                                    if (tmp != null)
                                    {
                                        this.Story.TranslateWord(wnd.WebView2Ctrl);
                                    }
                                    else
                                    {
                                        this.Story.TranslatePhrase(wnd.WebView2Ctrl);
                                    }

                                    key_eve.Handled = true; // キー入力解除
                                }
                                break;
                            }
                    }
                }
            }
            catch (Exception ev)
            {
                ShowMessage.ShowErrorOK(ev.Message, "Error");
            }
        }
        #endregion

    }
}
