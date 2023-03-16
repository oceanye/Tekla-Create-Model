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
    public partial class MonitorCheck : Form
    {
        public MonitorCheck()
        {
            InitializeComponent();
        }

        private void MonitorCheck_Load(object sender, EventArgs e)
        {

        }

        public delegate void SetPos(int ipos, int count);

        public void SetTextMesssage(int ipos, int count)
        {
            if (this.InvokeRequired)
            {
                SetPos setpos = new SetPos(SetTextMesssage);

                this.Invoke(setpos, new object[] { ipos, count });
            }
            else
            {


                progressBar1.Maximum = count;
                progressBar1.Value = Convert.ToInt32(ipos);

                if (progressBar1.Value == progressBar1.Maximum)
                {

                    this.Close();

                }


            }


        }

    }
}
