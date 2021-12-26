using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace AsistenteDeEscritura
{
    class ProgressDisplay
    {
        public ProgressDisplay(int i_total)
        {
            m_stopWatch = new Stopwatch();
            m_stopWatch.Start();
            m_total = i_total;
        }
        public enum UpdateResult
        {
            ContinueProcessing,
            Quit
        }

        public UpdateResult UpdateProgress(int i_current)
        {
            long ellapsed = m_stopWatch.ElapsedMilliseconds;
            if(ellapsed > 1000 && m_ventana == null)
            {
                m_ventana = new VentanaTrabajando(m_total);
                m_ventana.Show();
            }

            if(m_ventana != null)
            {
                bool cancelled = m_ventana.UpdateProgress(i_current);
                if(cancelled)
                {
                    return UpdateResult.Quit;
                }
                else
                {
                    return UpdateResult.ContinueProcessing;
                }
            }
            else
            {
                return UpdateResult.ContinueProcessing;
            }
        }

        public void Finish()
        {
            if(m_ventana != null)
            {
                m_ventana.Close();
            }
        }

        private VentanaTrabajando m_ventana;
        private Stopwatch m_stopWatch;
        private int m_total;
    }
}
