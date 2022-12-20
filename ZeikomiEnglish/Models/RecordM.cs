using Microsoft.Win32;
using MVVMCore.BaseClass;
using MVVMCore.Common.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using ZeikomiEnglish.Common.Util;

namespace ZeikomiEnglish.Models
{
    public class RecordM : ModelBase
    {

        #region 音声の録音
        /// <summary>
        /// 合成音声の録音
        /// </summary>
        public static void RecordSingle(StoryM story, string voice_name, int rate)
        {
            // ダイアログのインスタンスを生成
            var dialog = new SaveFileDialog();

            // ファイルの種類を設定
            dialog.Filter = "テキストファイル (*.wav)|*.wav";

            // ダイアログを表示する
            if (dialog.ShowDialog() == true)
            {
                // nullチェック
                if (story.PhraseItems.SelectedItem == null)
                {
                    story.PhraseItems.SelectedItem = story.PhraseItems.First();
                }

                // 音声再生
                PhraseVoiceToFile(story.PhraseItems.SelectedItem!.Phrase, voice_name, rate, dialog.FileName);  // フレーズ再生
            }
        }
        #endregion

        #region 合成音声の録音
        /// <summary>
        /// 合成音声の録音
        /// </summary>
        public static void RecordVoiceMulti(StoryM story, string voice_name, int rate)
        {
            try
            {
                StringBuilder text = new StringBuilder();
                foreach (var tmp in story.PhraseItems.Items)
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
                    PhraseVoiceToFile(text.ToString(), voice_name, rate, dialog.FileName);  // フレーズ再生
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
        /// <param name="phrase">フレーズ</param>
        public static TimeSpan PhraseVoice(string phrase, string voiceinfo_name, int rate)
        {
            try
            {
                var synthesizer = new SpeechSynthesizer();
                var tmp = synthesizer.GetInstalledVoices();
                synthesizer.SetOutputToDefaultAudioDevice();
                synthesizer.SelectVoice(voiceinfo_name);
                synthesizer.Rate = rate;
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
        #region フレーズを音声再生する
        /// <summary>
        /// フレーズを音声再生する
        /// </summary>
        /// <param name="phrase">フレーズ</param>
        public static TimeSpan PhraseVoiceToFile(string phrase, string voiceinfo_name, int rate, string filepath)
        {
            try
            {
                var synthesizer = new SpeechSynthesizer();
                var tmp = synthesizer.GetInstalledVoices();
                synthesizer.SetOutputToDefaultAudioDevice();
                synthesizer.SelectVoice(voiceinfo_name);
                synthesizer.Rate = rate;
                Stopwatch sw = new Stopwatch();
                sw.Start();
                synthesizer.SetOutputToWaveFile(filepath);
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
