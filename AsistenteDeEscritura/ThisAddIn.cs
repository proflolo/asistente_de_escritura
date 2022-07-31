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
        Regex m_gerundio = new Regex("ndo(me|te|le|nos|s|les|se)?(lo|la|los|las)?$");
        Regex m_fonemas = new Regex("ll|rr|pr|pl|br|bl|fr|fl|tr|tl|dr|dl|cr|cl|gr|[b-df-hj-np-tv-z]");
        Regex m_alfaNumerico = new Regex("^[a-zA-Z0-9_]*$");
        HashSet<string> m_dicientes = new HashSet<string>(Constantes.k_dicientes);
        HashSet<string> m_adjetivos = new HashSet<string>(Constantes.k_adjetivos);
        List<string> m_prefijos = new List<string>(Constantes.k_prefijos);
        List<string> m_sufijos = new List<string>(Constantes.k_sufijos);
        HashSet<string> m_comunes = new HashSet<string>();
        //Word.WdColor k_flojoColor = Word.WdColor.wdColorOrange;
        //Word.WdColor k_fuerteColor = Word.WdColor.wdColorPlum;
        static int k_limiparTotal = 100;
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

        private StatsPanel m_statsWidget;
        private Microsoft.Office.Tools.CustomTaskPane m_statsPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            m_palabrasResaltadas = new List<Word.Range>();
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(OnBeforeSave);
            ComparadorDeMorfemas comparador = new ComparadorDeMorfemas();
            m_prefijos.Sort(comparador);
            m_sufijos.Sort(comparador);

            Lexemer lexemer = new Lexemer();

            foreach(string palabraComun in Constantes.k_comunes)
            {
                if (!palabraComun.Contains(" "))
                {
                    string lexemaComun = lexemer.ComputeLexema(palabraComun);
                    m_comunes.Add(palabraComun);
                    m_comunes.Add(lexemaComun);
                }
            }

            //Create an instance of the user control
            m_statsWidget = new StatsPanel();
            // Connect the user control and the custom task pane 
            m_statsPane = CustomTaskPanes.Add(m_statsWidget, "Estadísticas");
            m_statsPane.VisibleChanged += (object _sender, EventArgs _e) => { OnPanelVisibilityChanged(); };
            m_statsPane.Visible = false;
            AsistenteDeEscritura.InterceptKeys.SetHook();

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
            AsistenteDeEscritura.InterceptKeys.ReleaseHook();
        }
        enum FlagStrength
        {
            Fuerte,
            Flojo
        }

        enum LimpiarResult
        {
            ContinueProcessing,
            Quit
        }
        private LimpiarResult LimiparPalabrasResaltadas(Word.Range i_range, ProgressDisplay i_progress)
        {
            try
            {
                int numLimpiadas = 0;
                int total = i_range.Words.Count;
                foreach (Word.Range palabra in i_range.Words)
                {
                    if (palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavyHeavy)
                    {
                        palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    }
                    numLimpiadas++;
                    if(i_progress.UpdateProgress((numLimpiadas * k_limiparTotal) / (total)) == ProgressDisplay.UpdateResult.Quit)
                    {
                        return LimpiarResult.Quit;
                    }
                }
            }
            catch (System.Exception e)
            {
                return LimpiarResult.Quit;
            }
            return LimpiarResult.ContinueProcessing;
        }

        private void FlagRange(Word.Range i_range, FlagStrength i_fuerza, Word.WdColor i_color)
        {
            try
            {
                if (i_fuerza == FlagStrength.Fuerte)
                {
                    i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
                    i_range.Font.UnderlineColor = i_color;
                }
                else
                {
                    if (i_range.Font.Underline != Word.WdUnderline.wdUnderlineWavyHeavy)
                    {
                        i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
                        i_range.Font.UnderlineColor = i_color;
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
                i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
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
            const int k_repeticionesTotal = 100;
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_range.Words.Count + k_repeticionesTotal);
            LimiparPalabrasResaltadas(i_range, progressUpdater);
            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            int i = 0;
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
                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }
            int j = 0;
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
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousWord = word;
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i_range.Words.Count + j * k_repeticionesTotal / wordDictionary.Count) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                j++;
            }
            progressUpdater.Finish();
        }

        public void ResaltarRepeticionesLexemas()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltarRepeticionesLexemas(documentRange, true);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltarTodoCacofonia()
        {
            Word.Range documentRange = GetAroundParaghaphs();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltaCacofonia(documentRange);
            ResaltaRimas(documentRange, false);
            ResaltarRepeticionesLexemas(documentRange, false);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void ResaltarRepeticionesLexemas(Word.Range i_range, bool i_limpiar)
        {
            Lexemer lexemer = new Lexemer();
            const int k_repeticionesTotal = 100;
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_range.Words.Count + k_repeticionesTotal);
            if(i_limpiar)
            { 
                LimiparPalabrasResaltadas(i_range, progressUpdater);
            }
            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            int i = 0;
            foreach (Word.Range word in i_range.Words)
            {
                if (word.Text != null)
                {
                    string lex = lexemer.ComputeLexema(word.Text);
                    if (word.Text.Length > 4)
                    {
                        if (wordDictionary.ContainsKey(lex))
                        {
                            wordDictionary[lex].Add(word);
                        }
                        else
                        {
                            List<Word.Range> wordRanges = new List<Word.Range>();
                            wordRanges.Add(word);
                            wordDictionary.Add(lex, wordRanges);
                        }
                    }
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }
            int j = 0;
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
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorPink);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorPink);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorPink);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorPink);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousWord = word;
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i_range.Words.Count + j * k_repeticionesTotal / wordDictionary.Count) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                j++;
            }
            progressUpdater.Finish();
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

        Word.Range GetAroundParaghaphs()
        {
            Word.Selection selection = this.Application.Selection;
            if (selection != null && selection.Range != null)
            {
                int selectionStart = selection.Range.Start;
                int selectionEnd = selection.Range.End;
                foreach (Word.Paragraph paragraph in Globals.ThisAddIn.Application.ActiveDocument.Range().Paragraphs)
                {
                    int start = paragraph.Range.Start;
                    int end = paragraph.Range.End;
                    if(start <= selectionStart && end >= selectionEnd)
                    {

                        Word.Range currentP = paragraph.Range;
                        int min = currentP.Start;
                        int max = currentP.End;
                        if(paragraph.Previous() != null && paragraph.Previous().Range != null)
                        {
                            min = paragraph.Previous().Range.Start;
                        }

                        if (paragraph.Next() != null && paragraph.Next().Range != null)
                        {
                            max = paragraph.Next().Range.End;
                        }
                        return this.Application.ActiveDocument.Range(min, max);
                    }
                }
                
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
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_range.Sentences.Count);
            LimiparPalabrasResaltadas(i_range, progressUpdater);
            int sentenceIdx = 0;
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
                if(progressUpdater.UpdateProgress(k_limiparTotal + sentenceIdx) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                sentenceIdx++;
            }
            progressUpdater.Finish();
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
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal);
            LimiparPalabrasResaltadas(documentRange, progressUpdater);
            progressUpdater.Finish();
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltaRimas()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltaRimas(documentRange, true);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }


        private void ResaltaRimas(Word.Range i_range, bool i_limpiar)
        {
            const int k_rimasTotal = 100;
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_range.Words.Count + k_rimasTotal * 2);
            if(i_limpiar)
            {
                LimiparPalabrasResaltadas(i_range, progressUpdater);
            }
            Dictionary<string, List<Word.Range>> rimeConsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            Dictionary<string, List<Word.Range>> rimeAsonanteDictionary = new Dictionary<string, List<Word.Range>>();
            int i = 0;
            foreach (Word.Range word in i_range.Words)
            {
                if (word.Text != null)
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
                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
            }

            int j = 0;
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
                            FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                            FlagRange(previousWord, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }
                    previousWord = word;
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i_range.Words.Count + (j * k_rimasTotal) / rimeConsonanteDictionary.Count) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                j++;
            }

            j = 0;
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
                            FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                            FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }
                    previousWord = word;
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i_range.Words.Count + k_rimasTotal + (j * k_rimasTotal) / rimeConsonanteDictionary.Count) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                j++;
            }
            progressUpdater.Finish();
        }

        private void ResaltarAdvMente(Word.Range i_documentRange)
        {
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Words.Count);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);
            int i = 0;
            foreach (Word.Range word in i_documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                if (text.Length > 5)
                {
                    Match adverbioMente = m_adverbioMente.Match(text);
                    if (adverbioMente != null && adverbioMente.Success)
                    {
                        FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                    }
                }
                i++;
                if(progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
            }
            progressUpdater.Finish();
        }

        private void ResaltarGerundios(Word.Range i_documentRange)
        {
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Words.Count);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);
            int i = 0;
            foreach (Word.Range word in i_documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                if (text.Length > 5)
                {
                    Match gerundio = m_gerundio.Match(text);
                    if (gerundio != null && gerundio.Success)
                    {
                        FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                    }
                }

                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }
            progressUpdater.Finish();
        }

        public void ResaltarAdvMente()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltarAdvMente(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void ResaltarGerundios()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltarGerundios(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void CorregirGuiones(Word.Range i_documentRange)
        {
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Words.Count);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);
            const string k_guion = "—";
            int processed = 0;
            foreach (Word.Paragraph paragraph in i_documentRange.Paragraphs)
            {
                //string text = paragraph.Range.Text.Trim().ToLower();
                //int i = 0;
                //++i;
                bool isDialog = false;

                foreach(Word.Range word in paragraph.Range.Words)
                {
                    string text = word.Text.Trim().ToLower();
                    if(text.StartsWith("—"))
                    {
                        isDialog = true;
                        break;
                    }
                    else if (text.Length > 0)
                    {
                        isDialog = false;
                        break;
                    }
                }

                if (isDialog)
                {
                    uint guionCount = 0;
                    bool estoyEndialogo = false;
                    Word.Range previousWord = null;
                    Word.Range previous2Word = null;
                    bool rightAfterNarrativeGuion = false;
                    int i = 0;
                    foreach (Word.Range word in paragraph.Range.Words)
                    {
                        string rawText = word.Text;
                        string text = rawText.Trim().ToLower();

                        if(rightAfterNarrativeGuion)
                        {
                            string previous2Text = previous2Word.Text.Trim().ToLower();
                            bool endsInPunto = previous2Text.EndsWith(".");
                            bool endsInSigno = previous2Text.EndsWith("?") || previous2Text.EndsWith("!");
                            if(m_dicientes.Contains(text))
                            {
                                if(endsInPunto)
                                {
                                    //Error, los dicientes no pueden tener un punto al final del dialogo
                                    FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                                }
                            }
                            else
                            {
                                if(!endsInPunto && !endsInSigno)
                                {
                                    //Error, es No diciente y no termina en punto
                                    FlagRange(previousWord, FlagStrength.Flojo, Word.WdColor.wdColorPlum);
                                }
                            }
                        }
                        rightAfterNarrativeGuion = false;

                        bool isGuion = text.IndexOf(k_guion) != -1;
                        if (isGuion)
                        {
                            if (estoyEndialogo)
                            {
                                string lastText = previousWord.Text;
                                rightAfterNarrativeGuion = true;
                                if (!lastText.EndsWith(" "))
                                {
                                    //Error, falta espacio.
                                    FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                                }   
                            }
                            else
                            {
                                if (guionCount != 0)
                                {
                                    bool endsInPunto = text.EndsWith(".");
                                    if (!endsInPunto)
                                    {
                                        //Punto tras cerrar la acotacion narrativa
                                        FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorPlum);
                                    }
                                }
                            }
                            guionCount++;
                            estoyEndialogo = !estoyEndialogo;
                            
                        }
                       
                        previous2Word = previousWord;
                        previousWord = word;
                        if (progressUpdater.UpdateProgress(k_limiparTotal + i + processed) == ProgressDisplay.UpdateResult.Quit)
                        {
                            progressUpdater.Finish();
                            return;
                        }
                    }
                }
                processed += paragraph.Range.Words.Count;
            }
            progressUpdater.Finish();
        }


            public void CorregirGuiones()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            CorregirGuiones(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        private void ResaltarRaras(Word.Range i_documentRange)
        {

        }

        public void ResaltarRaras()
        {
            Word.Range documentRange = GetSelectedRange();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            ResaltarRaras(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public class WordInfo
        {
            public WordInfo(string i_word, bool i_rare)
            {
                words = new HashSet<string>();
                words.Add(i_word);
                usage = 1;
                isRare = i_rare;
                referenceWord = i_word;
            }
            public void Increment(string i_word)
            {
                usage++;
                words.Add(i_word);
                if(i_word.Length < referenceWord.Length)
                {
                    referenceWord = i_word;
                }
            }

            public uint usage;
            public string referenceWord;
            public HashSet<string> words;
            public bool isRare;
        }

        public Dictionary<string, WordInfo> ComputeWordUsage()
        {
            Dictionary<string, WordInfo> result = new Dictionary<string, WordInfo>();
            Lexemer lexemer = new Lexemer();

            Word.Range documentRange = GetSelectedRange();
            foreach(Word.Range word in documentRange.Words)
            {
                string text = word.Text.ToLower().Trim();
                Match match = m_alfaNumerico.Match(text);
                if (match != null && match.Success && text.Length > 0)
                {
                    string lex = lexemer.ComputeLexema(text);
                    if (result.ContainsKey(lex))
                    {
                        result[lex].Increment(text);
                    }
                    else
                    {
                        bool isRara = !m_comunes.Contains(lex) && !m_comunes.Contains(text);
                        result.Add(lex, new WordInfo(text, isRara));
                    }
                }
            }

            return result;
        }

        public class FraseInfo
        {
            public FraseInfo(Word.Range i_frase, int i_longitud, int i_atomos)
            {
                frase = i_frase;
                longitud = i_longitud;
                atomos = i_atomos;
            }

            public int longitud;
            public int atomos;
            public Word.Range frase;
           
        }

        public IList<FraseInfo> ComputeFraseUsage()
        {
            List<FraseInfo> result = new List<FraseInfo>();
            

            Word.Range documentRange = GetSelectedRange();
            foreach (Word.Range sentence in documentRange.Sentences)
            {
                int ritmo = 1;
                int numPalabras = sentence.Words.Count;
                foreach (Word.Range word in sentence.Words)
                {
                    Match match = m_rithmSeparatorExpression.Match(word.Text);
                    if (match != null && match.Success)
                    {
                        ritmo++;
                    }
                }
                FraseInfo info = new FraseInfo(sentence, numPalabras, ritmo);
                result.Add(info);
            }

            return result;
        }

        private void ResaltaDeLista(Word.Range i_documentRange, HashSet<string> i_list)
        {
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Words.Count);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);
            int i = 0;
            foreach (Word.Range word in i_documentRange.Words)
            {
                string text = word.Text.ToLower().Trim();
                if(i_list.Contains(text))
                {
                    FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                }

                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }
            progressUpdater.Finish();
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
            const int k_silabaTotal = 100;
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Words.Count + k_silabaTotal);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);
            long t1 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            Dictionary<string, List<SilabaInfo>> silabaDictionary = new Dictionary<string, List<SilabaInfo>>();
            int idx = 0;
            foreach (Word.Range word in i_documentRange.Words)
            {
                int silabaPosition = 0;
                if (word.Text != null)
                {
                    string text = RemoveAccents(word.Text.ToLower().Trim());
                    Match match = m_alfaNumerico.Match(text);
                    if (match != null && match.Success)
                    {
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
                    }
                }
                idx++;
                if (progressUpdater.UpdateProgress(k_limiparTotal + idx) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
            }
            long t2 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;

            int i = 0;
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
                            FlagRange(info.word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                            FlagRange(previousSilaba.Value.word, FlagStrength.Flojo, Word.WdColor.wdColorOrange);
                        }
                        else
                        {
                            //No nos afecta
                        }

                    }

                    previousSilaba = info;
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i_documentRange.Words.Count + (i * k_silabaTotal) / silabaDictionary.Count) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }

            progressUpdater.Finish();
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
            ProgressDisplay progressUpdater = new ProgressDisplay(k_limiparTotal + i_documentRange.Sentences.Count);
            LimiparPalabrasResaltadas(i_documentRange, progressUpdater);

            Dictionary<string, List<SilabaInfo>> silabaDictionary = new Dictionary<string, List<SilabaInfo>>();
            int longSentenceSize = 40;
            int shortSentenceSize = 30;
            int i = 0;
            foreach (Word.Range sentence in i_documentRange.Sentences)
            {
                if(sentence.Words.Count > longSentenceSize)
                {
                    foreach (Word.Range word in sentence.Words)
                    {
                        FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorYellow);
                    }
                }

                if (sentence.Words.Count > shortSentenceSize && sentence.Words.Count <= longSentenceSize)
                {
                    foreach (Word.Range word in sentence.Words)
                    {
                        FlagRange(word, FlagStrength.Flojo, Word.WdColor.wdColorYellow);
                    }
                }
                if (progressUpdater.UpdateProgress(k_limiparTotal + i) == ProgressDisplay.UpdateResult.Quit)
                {
                    progressUpdater.Finish();
                    return;
                }
                i++;
            }
            progressUpdater.Finish();
        }

        public void ResaltaFrasesLargas()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("repeticiones");
            Word.Range documentRange = GetSelectedRange();
            ResaltaFrasesLargas(documentRange);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public delegate void OnStatsClosed();
        bool m_visibilityChangeRequestedByUser = false;
        OnStatsClosed m_panelClosedCallback;
        public void MuestraEstadisticas(bool i_visible, OnStatsClosed i_callback)
        {
            m_panelClosedCallback = i_callback;
            m_visibilityChangeRequestedByUser = true;
            m_statsPane.Visible = i_visible;
            m_statsWidget.ActualizaEstaisticas();
            m_visibilityChangeRequestedByUser = false;
        }

        public void OnPanelVisibilityChanged()
        {
            if(!m_visibilityChangeRequestedByUser)
            {
                m_panelClosedCallback();
            }
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
