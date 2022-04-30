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

        Dictionary<int, ThisAddIn.FraseInfo> m_fraseInfos = new Dictionary<int, ThisAddIn.FraseInfo>();
        public void ActualizaEstaisticas()
        {
            statsView.Items.Clear();
            statsView.Groups[0].Items.Clear();
            statsView.Groups[1].Items.Clear();
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

            m_fraseInfos.Clear();
            ritmoView.Items.Clear();
            IList<ThisAddIn.FraseInfo> fraseUsage = Globals.ThisAddIn.ComputeFraseUsage();
            int i = 0;
            foreach (ThisAddIn.FraseInfo info in fraseUsage)
            {
                string text = info.frase.Text;
                string keyColumn = text;
                if(keyColumn.Length > 12)
                {
                    keyColumn = keyColumn.Substring(0, 11) + "...";
                }
                ListViewItem item = ritmoView.Items.Add(keyColumn);
                item.ToolTipText = info.frase.Text;
                ListViewItem.ListViewSubItem logintudItem = item.SubItems.Add(info.longitud.ToString());
                //═╞╡██
                if (info.longitud > 40)
                {
                    logintudItem.ForeColor = Color.Red;
                }
                else if (info.longitud > 30)
                {
                    logintudItem.ForeColor = Color.Orange;
                }
                string longitudStr = new string('█', info.atomos);

                ListViewItem.ListViewSubItem ritmoItem = item.SubItems.Add(longitudStr);
                //ritmoView.Groups[0].Items.Add(item);
                m_fraseInfos.Add(i, info);
                ++i;
            }
        }

        private void StatsPanel_Load(object sender, EventArgs e)
        {
            
        }

        private void ritmoView_DoubleClick(object sender, EventArgs e)
        {
            var selectedItems = ritmoView.SelectedItems;
            if(selectedItems.Count > 0)
            {
                ListViewItem selectedItem = selectedItems[0];
                if (m_fraseInfos.ContainsKey(selectedItem.Index))
                {
                    m_fraseInfos[selectedItem.Index].frase.Select();
                }
            }
        }
    }
}
