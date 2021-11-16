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

namespace AsistenteDeEscritura
{
    public partial class ThisAddIn
    {
        List<Word.Range> m_palabrasResaltadas;
        Regex m_rithmSeparatorExpression = new Regex("^[,;yo]\\s*$");
        Regex m_acento = new Regex("[áéíóú]");
        Regex m_aguda = new Regex("[aeiouns]$");
        Regex m_dipongo = new Regex("[aeiou]+");
        Regex m_consonante = new Regex("[b-df-hj-np-tv-z]");
        Regex m_adverbioMente = new Regex("mente$");
        HashSet<string> m_dicientes = new HashSet<string>(Constantes.k_dicientes);
        HashSet<string> m_adjetivos = new HashSet<string>(Constantes.k_adjetivos);
        List<string> m_prefijos = new List<string>(Constantes.k_prefijos);
        List<string> m_sufijos = new List<string>(Constantes.k_sufijos);

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
            this.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
            this.Application.UndoRecord.EndCustomRecord();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        enum FlagStrength
        {
            Fuerte,
            Flojo
        }
        private void LimiparPalabrasResaltadas()
        {
            try
            {
                Word.Range documentRange = this.Application.ActiveDocument.Range();
                foreach (Word.Range palabra in documentRange.Words)
                {
                    if (palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavy || palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavyHeavy)
                    {
                        palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    }
                }
                foreach (Word.Comment comment in this.Application.ActiveDocument.Comments)
                {
                    if (comment.Range.Text.StartsWith("<AE>"))
                    {
                        comment.DeleteRecursively();
                    }
                }
            }
            catch (System.Exception e)
            {

            }
        }

        private void FlagRange(Word.Range i_range, FlagStrength i_fuerza, Word.WdColor i_color)
        {
            string commentTex = "";
            if(i_fuerza == FlagStrength.Fuerte)
            {
                i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
                commentTex = "Repetición(Cercana): " + i_range.Text;
            }
            else
            {
                if(i_range.Font.Underline != Word.WdUnderline.wdUnderlineWavyHeavy)
                {
                    i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                    commentTex = "Repetición(Lejana): " + i_range.Text;
                }
            }
            //
            i_range.Font.UnderlineColor =i_color;
            //Word.Comment existingComment = null;
            //if(io_commentCache.ContainsKey(i_range.Start))
            //{
            //    existingComment = io_commentCache[i_range.Start];
            //}
            //
            //if(existingComment == null)
            //{
            //    Word.Comment newComment = this.Application.ActiveDocument.Comments.Add(i_range, "<AE>" + commentTex);
            //    io_commentCache.Add(i_range.Start, newComment);
            //}
            //else
            //{
            //    existingComment.Range.InsertAfter("\n" + commentTex);
            //}
            

            m_palabrasResaltadas.Add(i_range);
        }

        public void ResaltarRepeticiones()
        {
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();

            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            foreach(Word.Range word in documentRange.Words)
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
            Dictionary<int, Word.Comment> cacheDeComentarios = new Dictionary<int, Word.Comment>();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
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
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousWord = word;
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltarRitmo()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            foreach(Word.Range sentence in documentRange.Sentences)
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
                    FlagRange(word, FlagStrength.Flojo, color);
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
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
            LimiparPalabrasResaltadas();
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }



        public void ResaltaRimas()
        {
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();

            Dictionary<string, List<Word.Range>> rimeConsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            Dictionary<string, List<Word.Range>> rimeAsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            foreach (Word.Range word in documentRange.Words)
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
            Dictionary<int, Word.Comment> cacheDeComentarios = new Dictionary<int, Word.Comment>();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
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
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
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
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }
                    previousWord = word;
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltarMalsonantes()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            foreach (Word.Range word in documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                Match adverbioMente = m_adverbioMente.Match(text);
                if (adverbioMente != null && adverbioMente.Success)
                {
                    FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorLavender);
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void ResaltaDeLista(HashSet<string> i_list)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            foreach (Word.Range word in documentRange.Words)
            {
                string text = word.Text.ToLower().Trim();
                if(i_list.Contains(text))
                {
                    FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorLavender);
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltaDicientes()
        {
            ResaltaDeLista(m_dicientes);
        }

        public void ResaltaAdjetivos()
        {
            ResaltaDeLista(m_adjetivos);
        }

        public void ResaltaLexemasRepetidos()
        {
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();

            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            foreach (Word.Range word in documentRange.Words)
            {
                string text = word.Text.ToLower().Trim();
                if (text.Length > 3)
                {
                    string lexema = text;
                    //Buscamos el lexema de la palabra

                    foreach (string prefijo in m_prefijos)
                    {
                        if (lexema.StartsWith(prefijo))
                        {
                            lexema = lexema.Substring(prefijo.Length);
                            break;
                        }
                    }
                    bool changed = false;
                    do
                    {
                        changed = false;
                        foreach (string sufijo in m_sufijos)
                        {
                            if (lexema.EndsWith(sufijo) && lexema.Length > 0)
                            {
                                lexema = lexema.Substring(0, lexema.Length - sufijo.Length);
                                changed = true;
                                break;
                            }
                        }
                    } while (changed);

                    if (wordDictionary.ContainsKey(lexema))
                    {
                        wordDictionary[lexema].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        wordDictionary.Add(lexema, wordRanges);
                    }
                }
            }
            Dictionary<int, Word.Comment> cacheDeComentarios = new Dictionary<int, Word.Comment>();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            LimiparPalabrasResaltadas();
            foreach (var kv in wordDictionary)
            {
                string text = kv.Key;
                Word.Range previousWord = null;
                foreach (Word.Range word in kv.Value)
                {
                    if (previousWord != null)
                    {
                        if (word.Start - previousWord.Start < 50)
                        {
                            //Repetición cercana!
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousWord = word;
                }
            }
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
