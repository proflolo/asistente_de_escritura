using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Collections;
using System.Text.RegularExpressions;

namespace AsistenteDeEscritura
{
    public partial class ThisAddIn
    {
        string[] m_prefijos;
        string[] m_sufijos;


        class ComparadorDeMorfemas : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if(x.Length > y.Length)
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
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(OnBeforeSave);
            m_prefijos = new string[] { "a", "aero", "an", "ambi", "anfi", "ante", "anti", "arci", "archi", "arz", "auto", "bi", "bis", "biz", "bien", "biem", "co", "com", "con", "contra", "cuasi", "cuasi", "de", "des", "epi", "equi", "ex", "extra", "hemi", "hetero", "hiper", "hipo", "homo", "in", "im", "i", "im", "infra", "inter", "entre", "intra", "iso", "macro", "maxi", "mega", "micro", "mini", "mono", "multi", "neo", "peri", "pluri", "poli", "pos", "pre", "pro", "re", "retro", "semi", "seudo", "sub", "so", "super", "super", "sobre", "supra", "trans", "tras", "ultra", "uni", "vi", "vice", "viz"};
            m_sufijos = new string[] { "ada", "ado", "aje", "ción", "dicción", "ducción", "dura", "ección", "epción", "ido", "ión", "miento", "ncia", "ón", "scripción", "sición", "sión", "dad", "tad", "bilidad", "edad", "era", "ería", "ez", "eza", "ía", "idad", "ismo", "ncia", "ura", "dor", "dora", "dero", "dera", "ero", "era", "ista", "ado", "ario", "ia", "ero", "era", "eria", "able", "áceo", "aco", "al", "áneo", "ante", "ario", "ente", "iente", "ento", "érrimo", "ible", "ico", "ífico", "il", "ino", "ísimo", "ivo", "izo", "oso", "ear", "ecer", "ificar", "izar", "ar", "er", "ir", "o", "as", "a", "amos", "áis", "an", "o", "es", "e", "emos", "éis", "en", "o", "es", "e", "imos", "ís", "en", "e", "es", "e", "emos", "éis", "en", "e", "as", "a", "amos", "áis", "an", "a", "as", "a", "iamos", "áis", "an", "aba", "abas", "aba", "ábamos", "abais", "aban", "ía", "ías", "ía", "íamos", "íais", "ían", "ía", "ías", "ía", "íamos", "íais", "ían", "ara", "se", "ara", "ses", "ara", "se", "ára", "semos", "ara", "seis", "ara", "sen", "iera", "se", "iera", "ses", "iera", "se", "iera", "semos", "iera", "seis", "iera", "sen", "iera", "se", "iera", "ses", "iera", "se", "iera", "semos", "iera", "seis", "iera", "sen", "é", "aste", "ó", "amos", "asteis", "aron", "í", "iste", "ió", "imos", "isteis", "ieron", "í", "istes", "ió", "imos", "isteis", "ieron", "aré", "arás", "ará", "aremos", "aréis", "arán", "eré", "erás", "erá", "eremos", "eréis", "erán", "iré", "irás", "irá", "iremos", "iréis", "irán", "aría", "arías", "aría", "aríamos", "aríais", "arían", "ería", "erías", "ería", "eríamos", "eríais", "erían", "iría", "irías", "iría", "iríamos", "iríais", "irían", "a", "ad", "e", "ed", "e", "id", "ar", "ando", "ado", "a", "s", "er", "endo", "ido", "a", "s", "ir", "iendo", "ido", "a", "s"};

            //Ordenar de mayor a menor
            ComparadorDeMorfemas comparador = new ComparadorDeMorfemas();
            Array.Sort(m_prefijos, comparador);
            Array.Sort(m_sufijos, comparador);
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
            Word.Range documentRange = this.Application.ActiveDocument.Range();
            foreach(Word.Range palabra in documentRange.Words)
            {
                if(palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavy || palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavyHeavy)
                {
                    palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                }
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
            i_range.Font.UnderlineColor = i_color;
           
        }

        public void ResaltarRepeticiones()
        {
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();

            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            Dictionary<string, List<Word.Range>> lexemaDictionary = new Dictionary<string, List<Word.Range>>();
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
                    string lexema = text;
                    //Buscamos el lexema de la palabra
                    
                    foreach (string prefijo in m_prefijos)
                    {
                        if(text.StartsWith(prefijo))
                        {
                            lexema = lexema.Substring(prefijo.Length);
                            break;
                        }
                    }
                    foreach (string sufijo in m_sufijos)
                    {
                        if (text.EndsWith(sufijo) && lexema.Length > 0)
                        {
                            lexema = lexema.Substring(0, lexema.Length - sufijo.Length);
                            break;
                        }
                    }

                    if (lexemaDictionary.ContainsKey(lexema))
                    {
                        lexemaDictionary[lexema].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        lexemaDictionary.Add(lexema, wordRanges);
                    }
                }
            }
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

            foreach (var kv in lexemaDictionary)
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

        static Regex s_adverbioMente = new Regex("mente$");

        public void ResaltarMalsonantes()
        {
            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("asistenteEscritura");
            LimiparPalabrasResaltadas();
            foreach (Word.Range word in documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                Match match = s_adverbioMente.Match(text);
                if(match != null && match.Success)
                {
                    FlagRange(word, FlagStrength.Fuerte, Word.WdColor.wdColorOrange);
                }
            }

            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        static Regex s_vocalConAcento = new Regex("[áéíóú]");
        static Regex s_diptongos = new Regex("[aeiou]+");
        static Regex s_terminaEnVocalNS = new Regex("[aeiouns]$");


        string CalculaRima(string i_palabra)
        {
            Match matchAcento = s_vocalConAcento.Match(i_palabra);
            if(matchAcento != null && matchAcento.Success)
            {
                return i_palabra.Substring(matchAcento.Index);
            }
            else
            {
                MatchCollection diptongos = s_diptongos.Matches(i_palabra);
                Match terminaEnVocalNS = s_terminaEnVocalNS.Match(i_palabra);
                if(terminaEnVocalNS != null && terminaEnVocalNS.Success) //Es llana
                {
                    if(diptongos.Count >= 2)
                    {
                        return i_palabra.Substring(diptongos[diptongos.Count - 2].Index);
                    }
                    else
                    {
                        return i_palabra;
                    }
                }
                else //Es aguda
                {
                    return i_palabra.Substring(diptongos[diptongos.Count - 1].Index);
                }
                //Aguda -> NO termina en vocal, n, s (sinó estaría acentuada)
                //Llana -> termina en vocal, n, s (sinó estaría acentuada)
            }
        }

        public void ResaltarRimas()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("asistenteEscritura");
            LimiparPalabrasResaltadas();

            Word.Range documentRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            Dictionary<string, List<Word.Range>> wordDictionary = new Dictionary<string, List<Word.Range>>();
            foreach (Word.Range word in documentRange.Words)
            {
                string text = word.Text.Trim().ToLower();
                if (text.Length > 3)
                {
                    string rima = CalculaRima(text.ToLower());
                    if (wordDictionary.ContainsKey(rima))
                    {
                        wordDictionary[rima].Add(word);
                    }
                    else
                    {
                        List<Word.Range> wordRanges = new List<Word.Range>();
                        wordRanges.Add(word);
                        wordDictionary.Add(rima, wordRanges);
                    }
                }

            }

            foreach (var kv in wordDictionary)
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

        public void Limpiar()
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("asistenteEscritura");
            LimiparPalabrasResaltadas();
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
