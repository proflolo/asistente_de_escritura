using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AsistenteDeEscritura
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarRepeticiones();
        }

        private void Malsonante_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarMalsonantes();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResaltarRimas();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Limpiar();

        }
    }
}
