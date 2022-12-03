using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;

namespace ZeikomiEnglish.Common.Util
{
    public class VoiceUtility
    {

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
    }
}
