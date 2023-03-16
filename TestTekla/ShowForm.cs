using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestTekla
{
    public partial class ShowForm : Form
    {
        public ShowForm(List<string> plist)
        {
            PList = plist;
            InitializeComponent();
        }


        public List<string> PList;

        private void ShowForm_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < PList.Count; i++)
            {
                listB.Items.Add(PList[i]);
            }
        }
    }
}
