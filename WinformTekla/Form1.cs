using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using Tekla.Structures;
using System.Data.SQLite;

using Tekla.Structures.Catalogs;
using Tekla.Structures.Drawing;

namespace WinformTekla
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void btn_Click(object sender, EventArgs e)
        {
            DrawingHandler MyDrawingHandler = new DrawingHandler();

            Drawing currentDraw = MyDrawingHandler.GetDrawings();

            DrawingObjectEnumerator DOE= currentDraw.GetSheet().GetAllObjects();
            while (DOE.MoveNext())
            {
                Tekla.Structures.Drawing.Polygon ply = DOE.Current as Tekla.Structures.Drawing.Polygon;
                if (ply != null)
                {
                    PointList plist = ply.Points;
                }

            }


        }
    }
}
