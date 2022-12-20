using Microsoft.Web.WebView2.WinForms;
using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using MVVMCore.Common.Wrapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Models
{
    public class StoryM : ModelBase
    {

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
        #region フレーズを翻訳した回数[PhraseTranslateCount]プロパティ
        /// <summary>
        /// フレーズを翻訳した回数[PhraseTranslateCount]プロパティ用変数
        /// </summary>
        int _PhraseTranslateCount = 0;
        /// <summary>
        /// フレーズを翻訳した回数[PhraseTranslateCount]プロパティ
        /// </summary>
        public int PhraseTranslateCount
        {
            get
            {
                return _PhraseTranslateCount;
            }
            set
            {
                if (!_PhraseTranslateCount.Equals(value))
                {
                    _PhraseTranslateCount = value;
                    NotifyPropertyChanged("PhraseTranslateCount");
                }
            }
        }
        #endregion

        #region 単語を翻訳した回数[WordTranslateCount]プロパティ
        /// <summary>
        /// 単語を翻訳した回数[WordTranslateCount]プロパティ用変数
        /// </summary>
        int _WordTranslateCount = 0;
        /// <summary>
        /// 単語を翻訳した回数[WordTranslateCount]プロパティ
        /// </summary>
        public int WordTranslateCount
        {
            get
            {
                return _WordTranslateCount;
            }
            set
            {
                if (!_WordTranslateCount.Equals(value))
                {
                    _WordTranslateCount = value;
                    NotifyPropertyChanged("WordTranslateCount");
                }
            }
        }
        #endregion



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

        #region 合成音声の初期化処理
        /// <summary>
        /// 合成音声の初期化処理
        /// </summary>
        public void InitVoice()
        {
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
        #endregion

        #region フレーズを音声再生する
        /// <summary>
        /// フレーズを音声再生する
        /// </summary>
        public void PhraseVoiceSingle()
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
                    var tm = RecordM.PhraseVoice(this.PhraseItems.SelectedItem!.Phrase, this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate);  // フレーズ再生

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
        #endregion

        #region 合成音声を連続で再生する
        /// <summary>
        /// 合成音声を連続で再生する
        /// </summary>
        public void PhraseVoiceMulti()
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
                        var tm = RecordM.PhraseVoice(tmp.Phrase, this.VoiceList.SelectedItem.VoiceInfo.Name, this.SpeechRate);    // フレーズ再生

                        if (this.PhraseItems.Count > index && this.PhraseItems.ElementAt(index) != null)
                        {
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
        #endregion


        #region フレーズの翻訳
        /// <summary>
        /// フレーズの翻訳
        /// </summary>
        /// <param name="wv2">WebView2コントロール</param>
        public void TranslatePhrase(Microsoft.Web.WebView2.Wpf.WebView2 wv2)
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
                if (wv2 != null && wv2.CoreWebView2 != null)
                {
                    // URLを開く
                    wv2.CoreWebView2.Navigate(url);

                    // フレーズ検索回数
                    this.PhraseTranslateCount++;

                    // 個別フレーズ翻訳回数のインクリメント
                    this.PhraseItems.SelectedItem.TranslateCount++;
                }
            }
        }
        #endregion

        #region 単語の検索
        /// <summary>
        /// 単語の検索
        /// </summary>
        /// <param name="wv2">WebView2コントロール</param>
        public void TranslateWord(Microsoft.Web.WebView2.Wpf.WebView2 wv2)
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
                if (wv2 != null && wv2.CoreWebView2 != null)
                {
                    // URLを開く
                    wv2.CoreWebView2.Navigate(url);

                    // 単語検索回数インクリメント
                    this.WordTranslateCount++;

                    // 個別単語翻訳回数のインクリメント
                    this.PhraseItems.SelectedItem.Words.SelectedItem.TranslateCount++;
                }
            }
        }
        #endregion

        #region フレーズリストの更新
        /// <summary>
        /// フレーズリストの更新
        /// </summary>
        public void RefreshPhraseList()
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

        #region 文章の整形処理
        /// <summary>
        /// 文章の整形処理
        /// </summary>
        public void PeriodLineBreak()
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
                    .Replace(";", ";\r\n");
            }
        }
        #endregion
    }
}
