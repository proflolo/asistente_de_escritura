using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

namespace AsistenteDeEscritura
{
    public partial class StatsPanel : UserControl
    {
        public StatsPanel()
        {
            InitializeComponent();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        class ComparadorDeInfo : IComparer<ThisAddIn.WordInfo>
        {
            public int Compare(ThisAddIn.WordInfo x, ThisAddIn.WordInfo y)
            {
                if (x.usage > y.usage)
                {
                    return -1;
                }
                else if (x.usage < y.usage)
                {
                    return 1;
                }
                else
                {
                    CaseInsensitiveComparer comparer = new CaseInsensitiveComparer();
                    return comparer.Compare(x.referenceWord, y.referenceWord);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ActualizaEstaisticas();
        }

        public void ActualizaEstaisticas()
        {
            statsView.Items.Clear();
            Dictionary<string, ThisAddIn.WordInfo> usage = Globals.ThisAddIn.ComputeWordUsage();
            List<ThisAddIn.WordInfo> sortedInfos = usage.Values.ToList();
            sortedInfos.Sort(new ComparadorDeInfo());
            foreach (ThisAddIn.WordInfo info in sortedInfos)
            {
                if (info.usage > 1)
                {
                    ListViewItem item = statsView.Items.Add(info.referenceWord + " x" + info.usage.ToString());
                    item.ToolTipText = String.Join(", ", info.words);
                    if (info.isRare)
                    {
                        statsView.Groups[0].Items.Add(item);
                    }
                    else
                    {
                        statsView.Groups[1].Items.Add(item);
                    }
                }
            }
        }

        private void StatsPanel_Load(object sender, EventArgs e)
        {
            
        }
    }
}
