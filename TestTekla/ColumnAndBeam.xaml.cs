using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using Tekla.Structures;
using System.Collections;

namespace TestTekla
{
    /// <summary>
    /// ColumnAndBeam.xaml 的交互逻辑
    /// </summary>
    public partial class ColumnAndBeam : Window
    {
        public ColumnAndBeam()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Model myModel = new Model();

            ArrayList objectToSelectBeam = new ArrayList();

            ArrayList objectToSelectColumn = new ArrayList();

           

            ModelObjectEnumerator myEnumBeam = myModel.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.BEAM);
          

            while (myEnumBeam.MoveNext())
            {
                if((myEnumBeam.Current as Beam).Name=="BEAM")
                {
                    objectToSelectBeam.Add(myEnumBeam.Current as Beam);
                }
                else 
                {
                    objectToSelectColumn.Add(myEnumBeam.Current as Beam);
                }
                
            }

            for (int i = 0; i < objectToSelectColumn.Count; i++)
            {



                Beam column = objectToSelectColumn[i] as Beam;




                ArrayList columnCenterLinePts = column.GetCenterLine(false);
                Tekla.Structures.Geometry3d.Line columnCenter = new Tekla.Structures.Geometry3d.Line(columnCenterLinePts[0] as Tekla.Structures.Geometry3d.Point, columnCenterLinePts[1] as Tekla.Structures.Geometry3d.Point);

                List<Beam> beamList = new List<Beam>();

                for (int j = 0; j < objectToSelectBeam.Count; j++)
                {

                    Beam baem = objectToSelectBeam[j] as Beam;



                    ArrayList beamCenterLinePts = baem.GetCenterLine(false);
                    Tekla.Structures.Geometry3d.Line beamCenter = new Tekla.Structures.Geometry3d.Line(beamCenterLinePts[0] as Tekla.Structures.Geometry3d.Point, beamCenterLinePts[1] as Tekla.Structures.Geometry3d.Point);


                    Tekla.Structures.Geometry3d.LineSegment lineSegment = Intersection.LineToLine(columnCenter, beamCenter);

                    if (Distance.PointToPoint(lineSegment.Point1, lineSegment.Point2) == 0)
                    {
                        beamList.Add(baem);
                    }

                }
            }



        }
    }
}
