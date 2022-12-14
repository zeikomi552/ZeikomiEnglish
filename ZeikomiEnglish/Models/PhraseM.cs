using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Models
{
    public class PhraseM : ModelBase
    {
        #region フレーズ[Phrase]プロパティ
        /// <summary>
        /// フレーズ[Phrase]プロパティ用変数
        /// </summary>
        string _Phrase = string.Empty;
        /// <summary>
        /// フレーズ[Phrase]プロパティ
        /// </summary>
        public string Phrase
        {
            get
            {
                return _Phrase;
            }
            set
            {
                if (_Phrase == null || !_Phrase.Equals(value))
                {
                    _Phrase = value;
                    NotifyPropertyChanged("Phrase");
                    WordSplit();
                }
            }
        }
        #endregion

        #region 単語[Words]プロパティ
        /// <summary>
        /// 単語[Words]プロパティ用変数
        /// </summary>
        ModelList<WordM> _Words = new ModelList<WordM>();
        /// <summary>
        /// 単語[Words]プロパティ
        /// </summary>
        public ModelList<WordM> Words
        {
            get
            {
                return _Words;
            }
            set
            {
                if (_Words == null || !_Words.Equals(value))
                {
                    _Words = value;
                    NotifyPropertyChanged("Words");
                    NotifyPropertyChanged("WordCount");
                }
            }
        }
        #endregion

        #region 翻訳した回数[TranslateCount]プロパティ
        /// <summary>
        /// 翻訳した回数[TranslateCount]プロパティ用変数
        /// </summary>
        int _TranslateCount = 0;
        /// <summary>
        /// 翻訳した回数[TranslateCount]プロパティ
        /// </summary>
        public int TranslateCount
        {
            get
            {
                return _TranslateCount;
            }
            set
            {
                if (!_TranslateCount.Equals(value))
                {
                    _TranslateCount = value;
                    NotifyPropertyChanged("TranslateCount");
                }
            }
        }
        #endregion

        #region 単語数
        /// <summary>
        /// 単語数
        /// </summary>
        public int PlayBackWordCount
        {
            get
            {
                return this.Words.Items.Count;
            }
        }
        #endregion

        #region 再生回数[PlayCount]プロパティ
        /// <summary>
        /// 再生回数[PlayCount]プロパティ用変数
        /// </summary>
        int _PlayCount = 0;
        /// <summary>
        /// 再生回数[PlayCount]プロパティ
        /// </summary>
        public int PlayCount
        {
            get
            {
                return _PlayCount;
            }
            set
            {
                if (!_PlayCount.Equals(value))
                {
                    _PlayCount = value;
                    NotifyPropertyChanged("PlayCount");
                }
            }
        }
        #endregion

        #region スピーチの再生時間[SpeechSec]プロパティ
        /// <summary>
        /// スピーチの再生時間[SpeechSec]プロパティ用変数
        /// </summary>
        double _SpeechSec = 0.0;
        /// <summary>
        /// スピーチの再生時間[SpeechSec]プロパティ
        /// </summary>
        public double SpeechSec
        {
            get
            {
                return _SpeechSec;
            }
            set
            {
                if (!_SpeechSec.Equals(value))
                {
                    _SpeechSec = value;
                    NotifyPropertyChanged("SpeechSec");
                }
            }
        }
        #endregion

        #region 単語で分割
        /// <summary>
        /// 単語で分割
        /// </summary>
        public void WordSplit()
        {
            this.Words.Items.Clear();
            var tmp = this.Phrase.Replace(",", " ")
                    .Replace("\\", " ")
                    .Replace(".", " ")
                    .Replace("\"", " ")
                    .Replace(";", " ")
                    .Replace(":", " ")
                    .Replace("?", " ")
                    .Split(" ");

            foreach (var word in tmp)
            {
                // 空文字なら飛ばす
                if (string.IsNullOrEmpty(word.Trim()))
                {
                    continue;
                }

                // 単語要素に加える
                this.Words.Items.Add(new WordM()
                {
                    Word = word.Trim()
                }
                );
            }
        }
        #endregion
    }
}
