using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace AsistenteDeEscritura
{
    public partial class ThisAddIn
    {
        List<Word.Range> m_palabrasResaltadas;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            m_palabrasResaltadas = new List<Word.Range>();
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(OnBeforeSave);

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
            //foreach(Word.Range palabra in m_palabrasResaltadas)
            //{
            //    palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            //}

            //Word.Range documentRange = this.Application.ActiveDocument.Range();
            //foreach(Word.Range palabra in documentRange.Words)
            //{
            //    if(palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavy || palabra.Font.Underline == Word.WdUnderline.wdUnderlineWavyHeavy)
            //    {
            //        palabra.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            //    }
            //}
            foreach(Word.Comment comment in this.Application.ActiveDocument.Comments)
            {
                if(comment.Range.Text.StartsWith("<AE>"))
                {
                    comment.DeleteRecursively();
                }
            }
        }

        private void FlagRange(Word.Range i_range, FlagStrength i_fuerza, IDictionary<int, Word.Comment> io_commentCache)
        {
            string commentTex = "";
            if(i_fuerza == FlagStrength.Fuerte)
            {
                //i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavyHeavy;
                commentTex = "Repetición(Cercana): " + i_range.Text;
            }
            else
            {
                if(i_range.Font.Underline != Word.WdUnderline.wdUnderlineWavyHeavy)
                {
                    //i_range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                    commentTex = "Repetición(Lejana): " + i_range.Text;
                }
            }
            //
            //i_range.Font.UnderlineColor = Word.WdColor.wdColorOrange;
            Word.Comment existingComment = null;
            if(io_commentCache.ContainsKey(i_range.Start))
            {
                existingComment = io_commentCache[i_range.Start];
            }
            
            if(existingComment == null)
            {
                Word.Comment newComment = this.Application.ActiveDocument.Comments.Add(i_range, "<AE>" + commentTex);
                io_commentCache.Add(i_range.Start, newComment);
            }
            else
            {
                existingComment.Range.InsertAfter("\n" + commentTex);
            }
            

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
                            FlagRange(word, FlagStrength.Fuerte, cacheDeComentarios);
                            FlagRange(previousWord, FlagStrength.Fuerte, cacheDeComentarios);
                        }
                        else if (word.Start - previousWord.Start < 100)
                        {
                            //repeticion lejana
                            FlagRange(word, FlagStrength.Flojo, cacheDeComentarios);
                            FlagRange(previousWord, FlagStrength.Flojo, cacheDeComentarios);
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
