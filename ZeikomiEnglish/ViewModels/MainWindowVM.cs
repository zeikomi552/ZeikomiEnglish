using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using MVVMCore.Common.Wrapper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
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

        public override void Init(object sender, EventArgs e)
        {
            try
            {
                var wnd = sender as MainWindow;
                if (wnd != null)
                {
                    wnd.WebView2Ctrl.EnsureCoreWebView2Async(null);
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
                    this.VoiceList.SelectedItem = this.VoiceList.Items.FirstOrDefault()!;
                }

            }
            catch (Exception ex)
            {
                ShowMessage.ShowErrorOK(ex.Message, "Error");
            }
        }

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

        /// <summary>
        /// フレーズリストの選択が変化した場合の処理
        /// </summary>
        public void TextChanged()
        {
            // フレーズに分解する
            var list = this.Text.Replace("\r", "").Split("\n");

            // フレーズリストをクリア
            this.PhraseItems.Items.Clear();

            // リスト数文まわす
            foreach (var tmp in list)
            {
                if (tmp.Trim().Length > 0)
                {
                    // フレーズリストに追加
                    this.PhraseItems.Items.Add(new PhraseM()
                    {
                        Phrase = tmp
                    }
                    );
                }
            }
        }


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
                    var deepl_url_base = "https://www.deepl.com/ja/translator#en/ja/{0}";   // DeepLのURL

                    // nullチェック
                    if (this.PhraseItems.SelectedItem != null)
                    {
                        string url = string.Format(deepl_url_base, this.PhraseItems.SelectedItem.Phrase);   // URL作成

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
                    var googletranslate_url_base = "https://translate.google.co.jp/?sl=en&tl=ja&text={0}&op=translate";

                    // nullチェック
                    if (this.PhraseItems.SelectedItem != null && this.PhraseItems.SelectedItem.Words.SelectedItem != null)
                    {
                        string url = string.Format(googletranslate_url_base, this.PhraseItems.SelectedItem.Words.SelectedItem.Word);   // URL作成

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



        public void PeriodLineBreak()
        {
            try
            {
                if (this.Text != null)
                {
                    this.Text = this.Text.Replace("\r", "").Replace("\n", "").Replace(".", ".\r\n").Replace(".\r\n\"", ".\"\r\n");
                }
            }
            catch
            {

            }
        }


        public void SelectedItemChanged(object sender, EventArgs e)
        {
            try
            {
                // ウィンドウを取得
                var wnd = VisualTreeHelperWrapper.GetWindow<MainWindow>(sender) as MainWindow;

                if (wnd != null)
                {
                    wnd.phrase_dg.ScrollIntoView(wnd.phrase_dg.Items[wnd.phrase_dg.SelectedIndex]); //scroll to last
                    wnd.phrase_dg.UpdateLayout();
                }
            }
            catch
            {

            }
        }
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
                        var tm = PhraseVoice(this.PhraseItems.SelectedItem!.Phrase);  // フレーズ再生
                        this.PhraseItems.ElementAt(index).SpeechSec = tm.TotalSeconds;

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
                    int index = 0;
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
                            var tm = PhraseVoice(tmp.Phrase);    // フレーズ再生

                            if (this.PhraseItems.Count > index && this.PhraseItems.ElementAt(index) != null)
                            {
                                this.PhraseItems.ElementAt(index).SpeechSec = tm.TotalSeconds;
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

        #region フレーズを音声再生する
        /// <summary>
        /// フレーズを音声再生する
        /// </summary>
        /// <param name="phrase">フレーズ</param>
        private TimeSpan PhraseVoice(string phrase)
        {
            try
            {
                var synthesizer = new SpeechSynthesizer();
                var tmp = synthesizer.GetInstalledVoices();
                synthesizer.SetOutputToDefaultAudioDevice();
                synthesizer.SelectVoice(this.VoiceList.SelectedItem.VoiceInfo.Name);
                synthesizer.Rate = this.SpeechRate;
                Stopwatch sw = new Stopwatch();
                sw.Start();
                synthesizer.Speak(phrase);
                sw.Stop();

                return sw.Elapsed;
            }
            catch
            {
                return TimeSpan.Zero;
            }
        }
        #endregion
    }
}
