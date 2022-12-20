using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
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
    }
}
