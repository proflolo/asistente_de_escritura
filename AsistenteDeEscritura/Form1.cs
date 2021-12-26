using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsistenteDeEscritura
{
    public partial class VentanaTrabajando : Form
    {
        bool m_cancelled = false;
        public VentanaTrabajando(int i_total)
        {
            InitializeComponent();
            this.progress.Maximum = i_total;
            this.totalLabel.Text = i_total.ToString();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            m_cancelled = true;
            Close();
        }

        private void totalLabel_Click(object sender, EventArgs e)
        {

        }

        public bool UpdateProgress(int i_current)
        {
            try
            {
                this.currentLabel.Text = i_current.ToString();
                this.progress.Value = i_current;
                Application.DoEvents();
                return m_cancelled;
            }
            catch (Exception e)
            {
                return m_cancelled;
            }
        }
    }
}
