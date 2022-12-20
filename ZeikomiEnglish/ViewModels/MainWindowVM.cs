﻿using ClosedXML.Excel;
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
                if (wnd != null)
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
    }
}
