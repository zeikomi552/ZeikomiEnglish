using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
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

        #region フレーズを調べた回数[PhraseSearch]プロパティ
        /// <summary>
        /// フレーズを調べた回数[PhraseSearch]プロパティ用変数
        /// </summary>
        int _PhraseSearch = 0;
        /// <summary>
        /// フレーズを調べた回数[PhraseSearch]プロパティ
        /// </summary>
        public int PhraseSearch
        {
            get
            {
                return _PhraseSearch;
            }
            set
            {
                if (!_PhraseSearch.Equals(value))
                {
                    _PhraseSearch = value;
                    NotifyPropertyChanged("PhraseSearch");
                }
            }
        }
        #endregion

        #region 単語を調べた回数[WordSearch]プロパティ
        /// <summary>
        /// 単語を調べた回数[WordSearch]プロパティ用変数
        /// </summary>
        int _WordSearch = 0;
        /// <summary>
        /// 単語を調べた回数[WordSearch]プロパティ
        /// </summary>
        public int WordSearch
        {
            get
            {
                return _WordSearch;
            }
            set
            {
                if (!_WordSearch.Equals(value))
                {
                    _WordSearch = value;
                    NotifyPropertyChanged("WordSearch");
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
            catch
            {
            }
        }
        #endregion
    }
}
