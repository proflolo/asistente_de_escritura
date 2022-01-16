using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace AsistenteDeEscritura
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarRepeticionesLexemas();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarRitmo();
        }

        private void Rimas_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltaRimas();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Limpiar();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarAdvMente();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltaDicientes();

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltaAdjetivos();

        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltaCacofonia();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            Globals.ThisAddIn.ResaltaCacofonia();
        }

        private void frases_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltaFrasesLargas();
        }

        private void Gerundios_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarGerundios();
        }

        private void Guiones_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CorregirGuiones();

        }

        private void RarasButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarRaras();
        }

        private void estadisticasButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.MuestraEstadisticas(estadisticasButton.Checked, () => { estadisticasButton.Checked = false;});
        }
    }
}
