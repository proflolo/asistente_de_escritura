using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Text.RegularExpressions;
using System.ComponentModel;

namespace AsistenteDeEscritura
{
    public partial class ThisAddIn
    {
        List<Word.Range> m_palabrasResaltadas;
        Regex m_rithmSeparatorExpression = new Regex("^[,;yo]\\s*$");
        Regex m_acento = new Regex("[áéíóú]");
        Regex m_aguda = new Regex("[aeiouns]$");
        Regex m_dipongo = new Regex("[aeiou]+");
        Regex m_vocalesDeSilaba = new Regex("(í|ú|[aeiouáéíóú]+)");
        Regex m_consonante = new Regex("[b-df-hj-np-tv-z]");
        Regex m_adverbioMente = new Regex("mente$");
        Regex m_fonemas = new Regex("ll|rr|pr|pl|br|bl|fr|fl|tr|tl|dr|dl|cr|cl|gr|[b-df-hj-np-tv-z]");
        HashSet<string> m_dicientes = new HashSet<string>(Constantes.k_dicientes);
        HashSet<string> m_adjetivos = new HashSet<string>(Constantes.k_adjetivos);
        List<string> m_prefijos = new List<string>(Constantes.k_prefijos);
        List<string> m_sufijos = new List<string>(Constantes.k_sufijos);
        Word.WdColor k_flojoColor = Word.WdColor.wdColorOrange;
        Word.WdColor k_fuerteColor = Word.WdColor.wdColorPlum;
        class ComparadorDeMorfemas : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (x.Length > y.Length)
                {
                    return -1;
                }
                else if (x.Length < y.Length)
                {
                    return 1;
                }
                else
                {
                    CaseInsensitiveComparer comparer = new CaseInsensitiveComparer();
                    return comparer.Compare(x, y);
                }
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            m_palabrasResaltadas = new List<Word.Range>();
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(OnBeforeSave);
            ComparadorDeMorfemas comparador = new ComparadorDeMorfemas();
            m_prefijos.Sort(comparador);
            m_sufijos.Sort(comparador);
        }

        private void OnBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //this.Application.UndoRecord.StartCustomRecord("repeticiones");
            //Word.Range documentRange = this.Application.ActiveDocument.Range();
            //LimiparPalabrasResaltadas(documentRange);
            //this.Application.UndoRecord.EndCustomRecord();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        enum FlagStrength
        {
            Fuerte,
            Flojo
        }
        private void LimiparPalabrasResaltadas(Word.Range i_range)
        {
            try
            {
                foreach (Word.Range palabra in i_range.Words)
                {
                    if (palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavy || palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavyHeavy)
                    {
                        palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    }
                }
            }
            catch (System.Exception e)
            {

            }
        }

        private void FlagRange(Word.Range i_range, FlagStrength i_fuerza)
        {
            try
            {
                if (i_fuerza == FlagStrength.Fuerte)
                {
                    i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
                    i_range.Font.UnderlineColor = k_fuerteColor;
                }
                else
                {
                    if (i_range.Font.Underline != Word.WdUnderline.wdUnderlineWavyHeavy)
                    {
                        i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                        i_range.Font.UnderlineColor = k_flojoColor;
                    }
                }
                //

                
            }
            catch (Exception e)
            {
                int i = 0;
                i++;
            }
            
            m_palabrasResaltadas.Add(i_range);
        }

        private void FlagRangewithColor(Word.Range i_range, Word.WdColor i_color)
        {
            try
            {
                i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                i_range.Font.UnderlineColor = i_color;
            }
            catch (Exception e)
            {
                int i = 0;
                i++;
            }

            m_palabrasResaltadas.Add(i_range);
        }

        public void ResaltarRepeticiones()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltarRepeticiones(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void ResaltarRepeticiones(Word.Range i_range)
        {

            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            foreach(Word.Range word in i_range.Words)
            {
                
                string text = word.Text.ToLower().Trim();
                if (text.Length > 3)
                {
                    if (wordDictionary.ContainsKey(text))
                    {
                        wordDictionary[text].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        wordDictionary.Add(text, wordRanges);
                    }
                }
            }
            LimiparPalabrasResaltadas(i_range);
            foreach (var kv in wordDictionary)
            {
                string text = kv.Key;
                Word.Range previousWord = null;
                foreach(Word.Range word in kv.Value)
                {
                    if(previousWord != null)
                    {
                        if (word.Start - previousWord.Start < 50)
                        {
                            //Repetición cercana!
                            FlagRange(word, FlagStrength.Fuerte);
                            FlagRange(previousWord, FlagStrength.Fuerte);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo);
                            FlagRange(previousWord, FlagStrength.Flojo);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousWord = word;
                }
            }
        }

        Word.Range GetSelectedRange()
        {
            Word.Selection selection = this.Application.Selection;
            if (selection != null && selection.Range != null && selection.Range.Text != null && selection.Range.Text.Length > 0)
            {
                return selection.Range;
            }
            else
            {
                return Globals.ThisAddIn.Application.ActiveDocument.Range();
            }
        }

        public void ResaltarRitmo()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltarRitmo(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void ResaltarRitmo(Word.Range i_range)
        {
            LimiparPalabrasResaltadas(i_range);
            foreach(Word.Range sentence in i_range.Sentences)
            {
                uint ritmo = 0;
                foreach(Word.Range word in sentence.Words)
                {
                    Match match = m_rithmSeparatorExpression.Match(word.Text);
                    if (match != null && match.Success)
                    {
                        ritmo++;
                    }
                }
                Word.WdColor color = Word.WdColor.wdColorBlack;
                switch(ritmo)
                {
                    case 0:
                        color = Word.WdColor.wdColorRed;
                        break;
                    case 1:
                        color = Word.WdColor.wdColorYellow;
                        break;
                    case 2:
                        color = Word.WdColor.wdColorGreen;
                        break;
                    case 3:
                        color = Word.WdColor.wdColorViolet;
                        break;
                    case 4:
                        color = Word.WdColor.wdColorBlue;
                        break;
                    default:
                        color = Word.WdColor.wdColorGray125;
                        break;
                }
                foreach (Word.Range word in sentence.Words)
                {
                    FlagRangewithColor(word, color);
                }
            }
        }

        private string ComputeRima(string i_text)
        {
            Match acento = m_acento.Match(i_text);
            if(acento != null && acento.Success)
            {
                return i_text.Substring(acento.Index);
            }
            else
            {
                MatchCollection matchCollection = m_dipongo.Matches(i_text);
                //Si fuese aguda y termina en vocal, N o S, entonces tendría acento. Pero no tiene, así que si termina en vocal, n o s, es llana
                Match agudaMatch = m_aguda.Match(i_text);
                if(agudaMatch != null && agudaMatch.Success)
                {
                    if(matchCollection.Count >= 2)
                    {
                        return i_text.Substring(matchCollection[matchCollection.Count - 2].Index);
                    }
                    else
                    {
                        return i_text;
                    }
                }
                else
                {
                    if (matchCollection.Count >= 2)
                    {
                        return i_text.Substring(matchCollection[matchCollection.Count - 1].Index);
                    }
                    else
                    {
                        return i_text;
                    }
                }
            }
        }

        public void Limpiar()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = this.Application.ActiveDocument.Range();
            LimiparPalabrasResaltadas(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltaRimas()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltaRimas(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }


        private void ResaltaRimas(Word.Range i_range)
        {
            Dictionary<string, List<Word.Range>> rimeConsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            Dictionary<string, List<Word.Range>> rimeAsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            foreach (Word.Range word in i_range.Words)
            {
                string raw = word.Text.ToLower().Trim();
                string rimaConsonante = ComputeRima(raw);
                string rimaAsonante = m_consonante.Replace(rimaConsonante, "_");
                if (raw.Length > 3)
                {
                    if (rimeConsonanteDictionary.ContainsKey(rimaConsonante))
                    {
                        rimeConsonanteDictionary[rimaConsonante].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        rimeConsonanteDictionary.Add(rimaConsonante, wordRanges);
                    }

                    if (rimeAsonanteDictionary.ContainsKey(rimaAsonante))
                    {
                        rimeAsonanteDictionary[rimaAsonante].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        rimeAsonanteDictionary.Add(rimaAsonante, wordRanges);
                    }
                }
            }
            
            LimiparPalabrasResaltadas(i_range);
            foreach (var kv in rimeConsonanteDictionary)
            {
                string text = kv.Key;
                Word.Range previousWord = null;
                foreach (Word.Range word in kv.Value)
                {
                    if (previousWord != null)
                    {
                        if (word.Start - previousWord.Start < 20)
                        {
                            //Repetición cercana!
                            FlagRange(word, FlagStrength.Fuerte);
                            FlagRange(previousWord, FlagStrength.Fuerte);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }
                    previousWord = word;
                }
            }
            foreach (var kv in rimeAsonanteDictionary)
            {
                string text = kv.Key;
                Word.Range previousWord = null;
                foreach (Word.Range word in kv.Value)
                {
                    if (previousWord != null)
                    {
                        if (word.Start - previousWord.Start < 20)
                        {
                            //Repetición cercana!
                            FlagRange(word, FlagStrength.Flojo);
                            FlagRange(previousWord, FlagStrength.Flojo);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }
                    previousWord = word;
                }
            }
        }

        private void ResaltarMalsonantes(Word.Range i_documentRange)
        {
            LimiparPalabrasResaltadas(i_documentRange);
            foreach (Word.Range word in i_documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                Match adverbioMente = m_adverbioMente.Match(text);
                if (adverbioMente != null && adverbioMente.Success)
                {
                    FlagRange(word, FlagStrength.Flojo);
                }
            }
        }

        public void ResaltarMalsonantes()
        {
            Word.Range documentRange = GetSelectedRange();
            ResaltarMalsonantes(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void ResaltaDeLista(Word.Range i_documentRange, HashSet<string> i_list)
        {
            LimiparPalabrasResaltadas(i_documentRange);
            foreach (Word.Range word in i_documentRange.Words)
            {
                string text = word.Text.ToLower().Trim();
                if(i_list.Contains(text))
                {
                    FlagRange(word, FlagStrength.Flojo);
                }
            }
        }

        public void ResaltaDicientes()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltaDeLista(documentRange, m_dicientes);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltaAdjetivos()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltaDeLista(documentRange, m_adjetivos);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

        }

        List<string> SeparaSilabas(string i_palabra)
        {
            MatchCollection matches = m_vocalesDeSilaba.Matches(i_palabra);
            if(matches.Count <= 1)
            {
                return new List<string>() { i_palabra };
            }

            List<string> silabas = new List<string>();
            string silaba = null;
            Match lastMatch = null;
            foreach (Match match in matches)
            {
                if(silaba == null)
                {
                    silaba = i_palabra.Substring(0, match.Index + match.Length);
                }
                else
                {
                    int lastEnd = lastMatch.Index + lastMatch.Length;
                    string entreDiptongos = i_palabra.Substring(lastEnd, match.Index - (lastMatch.Index + lastMatch.Length));
                    MatchCollection fonemas = m_fonemas.Matches(entreDiptongos);
                    int rightFonemas = Math.Max(1, fonemas.Count / 2);
                    if (fonemas.Count > 0)
                    {
                        int splitPoint = lastEnd + fonemas[fonemas.Count - (rightFonemas)].Index;
                        string izquierda = i_palabra.Substring(lastEnd, splitPoint - lastEnd);
                        string derecha = i_palabra.Substring(splitPoint, match.Index - splitPoint);
                        silaba += izquierda;
                        silabas.Add(silaba);
                        silaba = derecha + i_palabra.Substring(match.Index, match.Length);
                    }
                    else
                    {
                        silabas.Add(silaba);
                        silaba = i_palabra.Substring(match.Index, match.Length);
                    }
                }

                lastMatch = match;
            }

            if(silaba != null)
            {
                silabas.Add(silaba + i_palabra.Substring(lastMatch.Index + lastMatch.Length));
            }

            return silabas;
        }

        struct SilabaInfo
        {
            public Word.Range word;
            public int start;
            public int palabraIdx;
        };

        static Regex s_replace_a_Accents = new Regex("[á|à|ä|â]", RegexOptions.Compiled);
        static Regex s_replace_e_Accents = new Regex("[é|è|ë|ê]", RegexOptions.Compiled);
        static Regex s_replace_i_Accents = new Regex("[í|ì|ï|î]", RegexOptions.Compiled);
        static Regex s_replace_o_Accents = new Regex("[ó|ò|ö|ô]", RegexOptions.Compiled);
        static Regex s_replace_u_Accents = new Regex("[ú|ù|ü|û]", RegexOptions.Compiled);
        public static string RemoveAccents(string inputString)
        {
            inputString = s_replace_a_Accents.Replace(inputString, "a");
            inputString = s_replace_e_Accents.Replace(inputString, "e");
            inputString = s_replace_i_Accents.Replace(inputString, "i");
            inputString = s_replace_o_Accents.Replace(inputString, "o");
            inputString = s_replace_u_Accents.Replace(inputString, "u");
            return inputString;
        }

        public void ResaltaCacofonia(Word.Range i_documentRange)
        {
            long t1 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            Dictionary<string, List<SilabaInfo>> silabaDictionary = new Dictionary<string, List<SilabaInfo>>();
            int idx = 0;
            foreach (Word.Range word in i_documentRange.Words)
            {
                int silabaPosition = 0;
                string text = RemoveAccents(word.Text.ToLower().Trim());
                List<string> silabas = SeparaSilabas(text);
                foreach (string silaba in silabas)
                {
                    SilabaInfo info = new SilabaInfo();
                    info.word = word;
                    info.start = word.Start + silabaPosition;
                    info.palabraIdx = idx;
                    if (silabaDictionary.ContainsKey(silaba))
                    {
                        silabaDictionary[silaba].Add(info);
                    }
                    else
                    {
                        List<SilabaInfo> list = new List<SilabaInfo>();
                        list.Add(info);
                        silabaDictionary.Add(silaba, list);
                    }
                    silabaPosition += silaba.Length;
                }
                idx++;
            }
            long t2 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;

            //Agujero pelotero refresco aflojar calle cerro atlántico hambre inspector obstaculizar construcción instrucción malestar juan compadre casa pedro Tómate un té para que te alivies. – Toma un té, sentirás alivio. Tú que estuviste allí ¿viste lo que sucedió? – ¿Presenciaste lo que allí sucedió ? Ella me preguntó que qué estaba haciendo. – Ella me preguntó qué estaba haciendo. Colócalo donde coloqué los libros de cocina. – Colócalo donde están los libros de cocina. Las ballenas me llenan de alegría. – Las ballenas me dan alegría.
            LimiparPalabrasResaltadas(i_documentRange);
            foreach (var kv in silabaDictionary)
            {
                string text = kv.Key;
                SilabaInfo? previousSilaba = null;
                foreach (SilabaInfo info in kv.Value)
                {
                    if (previousSilaba != null)
                    {
                        if (info.start - previousSilaba.Value.start < 14 && info.palabraIdx != previousSilaba.Value.palabraIdx)
                        {
                            //Repetición cercana!
                            FlagRange(info.word, FlagStrength.Flojo);
                            FlagRange(previousSilaba.Value.word, FlagStrength.Flojo);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousSilaba = info;
                }
            }

            long t3 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            long p1 = t2 - t1;
            long p2 = t3 - t2;
        }
        public void ResaltaCacofonia()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltaCacofonia(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltaFrasesLargas(Word.Range i_documentRange)
        {
            LimiparPalabrasResaltadas(i_documentRange);

            Dictionary<string, List<SilabaInfo>> silabaDictionary = new Dictionary<string, List<SilabaInfo>>();
            int longSentenceSize = 30;
            int shortSentenceSize = 20;
            foreach (Word.Range sentence in i_documentRange.Sentences)
            {
                if(sentence.Words.Count > longSentenceSize)
                {
                    foreach (Word.Range word in sentence.Words)
                    {
                        FlagRange(word, FlagStrength.Fuerte);
                    }
                }

                if (sentence.Words.Count > shortSentenceSize && sentence.Words.Count <= longSentenceSize)
                {
                    foreach (Word.Range word in sentence.Words)
                    {
                        FlagRange(word, FlagStrength.Flojo);
                    }
                }
            }
        }

        public void ResaltaFrasesLargas()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltaFrasesLargas(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }


        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
