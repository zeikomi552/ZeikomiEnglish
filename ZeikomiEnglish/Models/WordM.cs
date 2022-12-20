using MVVMCore.BaseClass;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Models
{
    public class WordM : ModelBase
    {
        #region 単語[Word]プロパティ
        /// <summary>
        /// 単語[Word]プロパティ用変数
        /// </summary>
        string _Word = string.Empty;
        /// <summary>
        /// 単語[Word]プロパティ
        /// </summary>
        public string Word
        {
            get
            {
                return _Word;
            }
            set
            {
                if (_Word == null || !_Word.Equals(value))
                {
                    _Word = value;
                    NotifyPropertyChanged("Word");
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
    }
}
