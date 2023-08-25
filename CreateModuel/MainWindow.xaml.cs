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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using Tekla.Structures;
using System.Data.SQLite;
using System.Data;
using Tekla.Structures.Catalogs;
using System.Collections;
using System.IO;
using System.Windows.Forms;

namespace TestTekla
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


       public string DbPath = @"C:\ProgramData\Autodesk\Revit\Addins\2018";

        


       List<string> ProfileList = new List<string>();

        List<string> PList = new List<string>();
        //  public string DbPath = @"F:\Tekla\RvtData_v0526_debug.db";



        private void button_Click(object sender, RoutedEventArgs e)
        {

            bool PIP_in_Centre = (bool)checkbox_PIP_CENTRE.IsChecked;


            Generate_Model();
            if(PIP_in_Centre)
            {
                Modify_PIP_position();
            }
            
        }


        public void Generate_Model()
        {
            string f_path = DbPath;
            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018";

            }

            using (SQLiteConnection conn = new SQLiteConnection("Data Source = " + f_path + @"\数据库\RevitData.db"))
            {
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                string sqlBeam = "SELECT*FROM SectPropertyBeam";

                DataTable profileBeamData = getData(sqlBeam, cmd);

                string sqlColumn = "SELECT*FROM SectPropertyColumn";

                DataTable profileColumData = getData(sqlColumn, cmd);


                string sqlTag = "SELECT*FROM Tag";

                DataTable tagData = getData(sqlTag, cmd);

                for (int x = 0; x < profileBeamData.Rows.Count; x++)
                {
                    ProfileList.Add(profileBeamData.Rows[x].ItemArray[1].ToString().Split(' ')[2].Replace('X', '*'));
                }
                for (int x = 0; x < profileColumData.Rows.Count; x++)
                {
                    ProfileList.Add(profileColumData.Rows[x].ItemArray[1].ToString().Split(' ')[2].Replace('X', '*'));
                }


                List<LibraryProfileItem> profileL = new List<LibraryProfileItem>();
                List<MaterialItem> materialL = new List<MaterialItem>();

                CatalogHandler CatalogHandler = new CatalogHandler();



                ProfileItemEnumerator ProfileItemEnumerator = CatalogHandler.GetLibraryProfileItems();
                MaterialItemEnumerator MaterialItemEn = CatalogHandler.GetMaterialItems();




                while (ProfileItemEnumerator.MoveNext())
                {
                    LibraryProfileItem LibraryProfileItem = ProfileItemEnumerator.Current as LibraryProfileItem;

                    profileL.Add(LibraryProfileItem);

                }

                while (MaterialItemEn.MoveNext())
                {
                    MaterialItem materialItem = MaterialItemEn.Current as MaterialItem;

                    materialL.Add(materialItem);

                }

                for (int i = 0; i < ProfileList.Count; i++)
                {
                    int Flag = 0;
                    for (int j = 0; j < profileL.Count; j++)
                    {
                        if (ProfileList[i].Contains(profileL[j].ProfileName))
                        {
                            Flag++;
                        }
                    }

                    if (Flag == 0)
                    {
                        PList.Add(ProfileList[i]);
                    }

                }

                if (PList.Count > 0)
                {
                    Model myModel = new Model();
                    string path = myModel.GetInfo().ModelPath + "\\";
                    CreateText(path, PList);
                    //ShowForm sf = new TestTekla.ShowForm(PList);
                    //sf.ShowDialog();

                    //  如果原有截面库不完全，提示重新导入截面
                    DialogResult result = System.Windows.Forms.MessageBox.Show("截面缺失，请重新导入工程文件下:截面文件.lis", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        Environment.Exit(0);

                    }


                }


                #region 删除原始轴网
                Model myModelGrid = new Model();

                List<Tekla.Structures.Model.Grid> objectToSelectGrid = new List<Tekla.Structures.Model.Grid>();


                ModelObjectEnumerator myEnumFitting = myModelGrid.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.GRID);


                while (myEnumFitting.MoveNext())
                {
                    objectToSelectGrid.Add(myEnumFitting.Current as Tekla.Structures.Model.Grid);
                }


                foreach (Tekla.Structures.Model.Grid g in objectToSelectGrid)
                {
                    g.Delete();
                }

                myModelGrid.CommitChanges();


                #endregion

                #region 生成新轴网
                cmd.CommandText = "SELECT*FROM LevelTable";

                SQLiteDataAdapter adapterGrid = new SQLiteDataAdapter(cmd);
                DataSet dsGrid = new DataSet();
                adapterGrid.Fill(dsGrid);

                DataTable tableGrid = dsGrid.Tables[0];
                Model modelNewGrid = new Model();

                Tekla.Structures.Model.Grid grid = new Tekla.Structures.Model.Grid();

                string Z = "";

                string LabelZ = "+0 ";

                for (int i = 0; i < tableGrid.Rows.Count; i++)
                {
                    Z += tableGrid.Rows[i].ItemArray[2].ToString() + " ";

                    LabelZ += tableGrid.Rows[i].ItemArray[1].ToString() + " ";
                }


                grid.Name = "Grid";
                grid.CoordinateX = "0.00";
                grid.CoordinateY = "0.00";
                grid.CoordinateZ = Z;
                grid.LabelX = "1";
                grid.LabelY = "A";
                grid.LabelZ = LabelZ;

                grid.ExtensionLeftX = 2000;
                grid.ExtensionLeftY = 2000;
                grid.ExtensionLeftZ = 2000;

                grid.ExtensionRightX = 2000;
                grid.ExtensionRightY = 2000;
                grid.ExtensionRightZ = 2000;

                grid.IsMagnetic = true;

                grid.Insert();

                modelNewGrid.CommitChanges();

                #endregion


                #region 创建梁

                cmd.CommandText = "SELECT*FROM MemberBeam";

                SQLiteDataAdapter adapterBeam = new SQLiteDataAdapter(cmd);
                DataSet dsBeam = new DataSet();
                adapterBeam.Fill(dsBeam);

                DataTable tableBeam = dsBeam.Tables[0];
                Model model = new Model();

                MonitorCheck form = new MonitorCheck();
                form.Show();
                int T = 0;

                int Rec_hole = 0;
                int Cir_hole = 0;

                for (int i = 0; i < tableBeam.Rows.Count; i++)
                {
                    T++;
                    form.SetTextMesssage(T, tableBeam.Rows.Count);//进度条函数


                    #region
                    Beam beam = new Beam();

                    string startStr = tableBeam.Rows[i].ItemArray[1].ToString().Substring(1, tableBeam.Rows[i].ItemArray[1].ToString().Length - 2);

                    string endStr = tableBeam.Rows[i].ItemArray[2].ToString().Substring(1, tableBeam.Rows[i].ItemArray[2].ToString().Length - 2);

                    Tekla.Structures.Geometry3d.Point SPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(startStr.Split(',')[0]), Convert.ToDouble(startStr.Split(',')[1]), Convert.ToDouble(startStr.Split(',')[2]));

                    Tekla.Structures.Geometry3d.Point EPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(endStr.Split(',')[0]), Convert.ToDouble(endStr.Split(',')[1]), Convert.ToDouble(endStr.Split(',')[2]));

                    string section_name = tableBeam.Rows[i].ItemArray[3].ToString();
                    string ProfileBeam = section_name.Split(' ')[2].Replace('X', '*');
                    string MaterialBeam = tableBeam.Rows[i].ItemArray[17].ToString();

                    Tekla.Structures.Geometry3d.Point VectorPoint = new Tekla.Structures.Geometry3d.Point(EPoint.X - SPoint.X, EPoint.Y - SPoint.Y, EPoint.Z - SPoint.Z);



                    if (ProfileBeam.StartsWith("T"))
                    {
                        Output_Text.AppendText("存在T型截面，需核对构件方向");
                    }
                    beam.StartPoint = GetNewStartPoint(SPoint, VectorPoint, Convert.ToDouble(tableBeam.Rows[i].ItemArray[11]));

                    beam.EndPoint = GetNewStartPoint(EPoint, VectorPoint, Convert.ToDouble(tableBeam.Rows[i].ItemArray[12]));



                    beam.Profile.ProfileString = "";
                    for (int x = 0; x < profileL.Count; x++)
                    {
                        if (profileL[x].ProfileName.Contains(ProfileBeam))
                        {
                            beam.Profile.ProfileString = ProfileBeam;
                        }
                    }

                    int flag = 0;
                    for (int x = 0; x < materialL.Count; x++)
                    {
                        if (materialL[x].MaterialName.Contains(MaterialBeam))
                        {
                            beam.Material.MaterialString = MaterialBeam;
                            flag++;
                        }
                    }
                    if (flag == 0)
                    {
                        beam.Material.MaterialString = "Q235B";
                    }



                    beam.Class = "3";


                    if (comboBox.SelectionBoxItem.ToString() == "是")
                    {
                        beam.StartPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[11]);
                        beam.EndPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[12]);
                        beam.StartPointOffset.Dz = Convert.ToDouble(tableBeam.Rows[i].ItemArray[15]);
                        beam.EndPointOffset.Dz = Convert.ToDouble(tableBeam.Rows[i].ItemArray[16]);
                    }





                    if (beam.Profile.ProfileString != "")
                    {
                        beam.Insert();


                    }


                    #region 测试


                    beam.SetUserProperty("comment", "1");
                    #endregion

                    #region 交点

                    CoordinateSystem beamLocalCS = beam.GetCoordinateSystem();

                    if (beamLocalCS != null)
                    {

                        GeometricPlane gPlan = new GeometricPlane(beamLocalCS);


                        Tekla.Structures.Geometry3d.Vector vector = gPlan.Normal;

                        for (int j = 0; j < tagData.Rows.Count; j++)
                        {
                            Tekla.Structures.Geometry3d.Point pt = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(tagData.Rows[j].ItemArray[1]), Convert.ToDouble(tagData.Rows[j].ItemArray[2]), Convert.ToDouble(tagData.Rows[j].ItemArray[3]));

                            Tekla.Structures.Geometry3d.Point ptNew = GetNewStartPointSecond(pt, vector, -1000);
                            Tekla.Structures.Geometry3d.Point ptNewSecond = GetNewStartPointSecond(pt, vector, 1000);



                            Tekla.Structures.Geometry3d.Line line = new Tekla.Structures.Geometry3d.Line(pt, ptNew);
                            Tekla.Structures.Geometry3d.Line lineSecond = new Tekla.Structures.Geometry3d.Line(pt, ptNewSecond);


                            Tekla.Structures.Geometry3d.Point pInter = Intersection.LineToPlane(line, gPlan);
                            Tekla.Structures.Geometry3d.Point pInterSecond = Intersection.LineToPlane(lineSecond, gPlan);

                        }
                    }

                    

                    #endregion

                    #endregion

                    #region 孔洞

                    string sqlHole = "SELECT*FROM BeamHoleTable";
                    DataTable holetable = getData(sqlHole, cmd);


                    for (int x = 0; x < holetable.Rows.Count; x++)
                    {
                        if (holetable.Rows[x].ItemArray[5].ToString().Equals(tableBeam.Rows[i].ItemArray[0].ToString()))
                        {

                            ContourPlate CP = new ContourPlate();

                            if (holetable.Rows[x].ItemArray[1].ToString().Contains("@"))// 存在@
                            {
                                // 中心点
                                string pc = holetable.Rows[x].ItemArray[1].ToString().Split('@')[1];
                                pc = pc.Substring(1, pc.Length - 2);

                                // 法线方向
                                string n1s = holetable.Rows[x].ItemArray[1].ToString().Split('@')[2];
                                n1s = n1s.Substring(1, n1s.Length - 2);

                                //半径
                                double rc = Convert.ToDouble(holetable.Rows[x].ItemArray[1].ToString().Split('@')[0]);
                                Tekla.Structures.Geometry3d.Point center = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(pc.Split(',')[0]), Convert.ToDouble(pc.Split(',')[1]), Convert.ToDouble(pc.Split(',')[2]) + beam.StartPointOffset.Dz);


                                double halfSideLength = rc / Math.Sqrt(2);

                                Tekla.Structures.Geometry3d.Vector n1 = new Tekla.Structures.Geometry3d.Vector(Convert.ToDouble(n1s.Split(',')[0]), Convert.ToDouble(n1s.Split(',')[1]), Convert.ToDouble(n1s.Split(',')[2]));


                                // 根据法向量计算矩形的四个点
                                // 根据法向量计算矩形的四个点
                                Tekla.Structures.Geometry3d.Vector rightDirection = n1.Cross(new Tekla.Structures.Geometry3d.Vector(0, 0, 1)).GetNormal();
                                Tekla.Structures.Geometry3d.Vector upDirection = n1.Cross(rightDirection).GetNormal();

                                Tekla.Structures.Geometry3d.Point p1 = center + rightDirection * halfSideLength - upDirection * halfSideLength;
                                Tekla.Structures.Geometry3d.Point p2 = center + rightDirection * halfSideLength + upDirection * halfSideLength;
                                Tekla.Structures.Geometry3d.Point p3 = center - rightDirection * halfSideLength + upDirection * halfSideLength;
                                Tekla.Structures.Geometry3d.Point p4 = center - rightDirection * halfSideLength - upDirection * halfSideLength;



                                // 创建切角对象
                                Chamfer chamfer = new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_ROUNDING);

                                // 在添加轮廓点时为其设置切角
                                CP.AddContourPoint(new ContourPoint(p4, chamfer));
                                CP.AddContourPoint(new ContourPoint(p3, chamfer));
                                CP.AddContourPoint(new ContourPoint(p2, chamfer));
                                CP.AddContourPoint(new ContourPoint(p1, chamfer));

                                Cir_hole = Cir_hole + 1;

                            }
                            else
                            {
                                string p1 = holetable.Rows[x].ItemArray[1].ToString().Substring(1, holetable.Rows[x].ItemArray[1].ToString().Length - 2);
                                string p2 = holetable.Rows[x].ItemArray[2].ToString().Substring(1, holetable.Rows[x].ItemArray[2].ToString().Length - 2);
                                string p3 = holetable.Rows[x].ItemArray[3].ToString().Substring(1, holetable.Rows[x].ItemArray[3].ToString().Length - 2);
                                string p4 = holetable.Rows[x].ItemArray[4].ToString().Substring(1, holetable.Rows[x].ItemArray[4].ToString().Length - 2);




                                Tekla.Structures.Geometry3d.Point Pt1 = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(p1.Split(',')[0]), Convert.ToDouble(p1.Split(',')[1]), Convert.ToDouble(p1.Split(',')[2]) + beam.StartPointOffset.Dz);
                                Tekla.Structures.Geometry3d.Point Pt2 = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(p2.Split(',')[0]), Convert.ToDouble(p2.Split(',')[1]), Convert.ToDouble(p2.Split(',')[2]) + beam.StartPointOffset.Dz);
                                Tekla.Structures.Geometry3d.Point Pt3 = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(p3.Split(',')[0]), Convert.ToDouble(p3.Split(',')[1]), Convert.ToDouble(p3.Split(',')[2]) + beam.StartPointOffset.Dz);
                                Tekla.Structures.Geometry3d.Point Pt4 = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(p4.Split(',')[0]), Convert.ToDouble(p4.Split(',')[1]), Convert.ToDouble(p4.Split(',')[2]) + beam.StartPointOffset.Dz);

                                if (Pt1 != Pt2)
                                {

                                    ContourPoint point = new ContourPoint(Pt1, null);
                                    ContourPoint point1 = new ContourPoint(Pt2, null);
                                    ContourPoint point2 = new ContourPoint(Pt3, null);
                                    ContourPoint point3 = new ContourPoint(Pt4, null);


                                    CP.AddContourPoint(point);
                                    CP.AddContourPoint(point1);
                                    CP.AddContourPoint(point2);
                                    CP.AddContourPoint(point3);

                                    Rec_hole = Rec_hole + 1;
                                }
                            }

                            CP.Finish = "FOO";
                            CP.Profile.ProfileString = "PL800";
                            CP.Material.MaterialString = "K30-2";
                            CP.Class = BooleanPart.BooleanOperativeClassName;


                            CP.Insert();
                            BooleanPart Beam = new BooleanPart();

                            Beam.Father = beam;

                            Beam.SetOperativePart(CP);

                            Beam.Insert();
                            CP.Delete();






                        }


                    }


                    


                    #endregion


                    model.CommitChanges();

                }

                Output_Text.AppendText("生成模型 梁:" + T + "个 \n");
                Output_Text.AppendText("包含梁上开洞"+(Rec_hole+Cir_hole)+"个：方" + Rec_hole + "/圆" + Cir_hole+"个 \n");


                #endregion

                #region 创建柱

                cmd.CommandText = "SELECT*FROM MemberColumn";

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                DataTable table = ds.Tables[0];

                Model modelCo = new Model();

                MonitorCheck form1 = new MonitorCheck();
                form1.Show();
                int T1 = 0;

                for (int i = 0; i < table.Rows.Count; i++)
                {

                    T1++;
                    form1.SetTextMesssage(T1, table.Rows.Count);//进度条函数




                    #region
                    Beam beam = new Beam();

                    string startStr = table.Rows[i].ItemArray[1].ToString().Substring(1, table.Rows[i].ItemArray[1].ToString().Length - 2);

                    string endStr = table.Rows[i].ItemArray[2].ToString().Substring(1, table.Rows[i].ItemArray[2].ToString().Length - 2);

                    string ProfileColumn_org = table.Rows[i].ItemArray[3].ToString().Split(' ')[2];
                    string ProfileColumn = ProfileColumn_org.Replace('X', '*');

                    string MaterialColumn = table.Rows[i].ItemArray[11].ToString();

                    beam.StartPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(startStr.Split(',')[0]), Convert.ToDouble(startStr.Split(',')[1]), Convert.ToDouble(startStr.Split(',')[2]));

                    beam.EndPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(endStr.Split(',')[0]), Convert.ToDouble(endStr.Split(',')[1]), Convert.ToDouble(endStr.Split(',')[2]));

                    beam.Profile.ProfileString = "";

                    for (int x = 0; x < profileL.Count; x++)
                    {
                        if (profileL[x].ProfileName.Contains(ProfileColumn))
                        {
                            beam.Profile.ProfileString = ProfileColumn;
                        }
                    }

                    beam.Position.Plane = Position.PlaneEnum.MIDDLE;

                    beam.Position.Depth = Position.DepthEnum.MIDDLE;


                    beam.Position.Rotation = Position.RotationEnum.TOP;


                    beam.Position.RotationOffset = Convert.ToDouble(table.Rows[i].ItemArray[10].ToString());


                    int flag = 0;
                    for (int x = 0; x < materialL.Count; x++)
                    {
                        if (materialL[x].MaterialName.Contains(MaterialColumn))
                        {
                            beam.Material.MaterialString = MaterialColumn;
                            flag++;
                        }
                    }
                    if (flag == 0)
                    {
                        beam.Material.MaterialString = "Q235B";
                    }


                    beam.Class = "7";
                    beam.Name = "column";


                    if (beam.Profile.ProfileString != "")
                    {
                        beam.Insert();
                    }

                    #endregion


                    modelCo.CommitChanges();
                }

                Output_Text.AppendText("生成模型 柱:" + T1 + "个");

                #endregion


                conn.Close();





            }







            //Close();




        }


        public void Modify_PIP_position()
        {

            try
            {
                Model model = new Model();

                ModelObjectEnumerator modelObjectEnumerator = new Model().GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.BEAM);
                while (modelObjectEnumerator.MoveNext())
                {
                    Beam b = modelObjectEnumerator.Current as Beam;
                    if (b != null)
                    {

                            if (b.Profile.ProfileString.StartsWith("P"))
                            {
                                b.Position.Depth = Position.DepthEnum.MIDDLE;
                                b.Modify();
                            }

                    }
                }

                model.CommitChanges();
                Console.WriteLine("Modified PIP sections successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

        }

        public void CreateText(string filePath,List<string> pList)
        {
            if (Directory.Exists(filePath) == false)
            {

                Directory.CreateDirectory(filePath);

            }

            
                #region 


                FileStream fil = new FileStream(filePath + "截面文件.lis", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);
                sw.WriteLine("PROFILE DATABASE EXPORT VERSION = 3");
                sw.WriteLine("");
                for (int i = 0; i < pList.Count; i++)
                {
                    string profileName = pList[i].Replace("X", "*");

                    char section_type = profileName[0];
                    if (section_type == 'H')
                      {
                        double  H =Convert.ToDouble( profileName.Split('*')[0].Substring(1));
                        double W =Convert.ToDouble( profileName.Split('*')[1]);
                        double F = Convert.ToDouble(profileName.Split('*')[2]);
                        double Y = Convert.ToDouble(profileName.Split('*')[3]);

                        sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\""+";");
                        sw.WriteLine("{");
                        sw.WriteLine("  TYPE = 1; SUB_TYPE = 1001; COORDINATE = 0.000;");
                        sw.WriteLine("  {");
                        sw.WriteLine("    \"HEIGHT\"" + "                           " + H );
                        sw.WriteLine("    \"WIDTH\"" + "                           " + W);
                        sw.WriteLine("    \"WEB_THICKNESS\"" + "                           " + F );
                        sw.WriteLine("    \"FLANGE_THICKNESS\"" + "                           " + Y);
                        sw.WriteLine("    \"ROUNDING_RADIUS_1\"" + "                           "+ "0.000000000E+000");
                        sw.WriteLine("    \"ROUNDING_RADIUS_2\"" + "                           " + "0.000000000E+000");
                        sw.WriteLine("    \"FLANGE_SLOPE_RATIO\"" + "                           " + "0.000000000E+000");
                        sw.WriteLine("  }");
                        sw.WriteLine("}");
                        sw.WriteLine("");
                       }
                    else if (section_type == 'B')
                        {
                            double H = Convert.ToDouble(profileName.Split('*')[0].Substring(1));
                            double W = Convert.ToDouble(profileName.Split('*')[1]);
                            double F = Convert.ToDouble(profileName.Split('*')[2]);
                            //double Y = Convert.ToDouble(profileName.Split('*')[3]);

                            sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                            sw.WriteLine("{");
                            sw.WriteLine("  TYPE = 8; SUB_TYPE = 8002; COORDINATE = 0.000;");
                            sw.WriteLine("  {");
                            sw.WriteLine("    \"HEIGHT\"" + "                           " + H);
                            sw.WriteLine("    \"WIDTH\"" + "                           " + W);
                            sw.WriteLine("    \"PLATE_THICKNESS\"" + "                           " + F);
                            sw.WriteLine("    \"ROUNDING_RADIUS\"" + "                           " + "1.000000000E+001");
                            sw.WriteLine("  }");
                            sw.WriteLine("}");
                            sw.WriteLine("");
                        }
                    else if (section_type == 'P')
                        {
                            double D = Convert.ToDouble(profileName.Split('*')[0].Substring(1));
                            double t = Convert.ToDouble(profileName.Split('*')[1]);


                            sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                            sw.WriteLine("{");
                            sw.WriteLine("  TYPE = 7; SUB_TYPE = 7001; COORDINATE = 0.000;");
                            sw.WriteLine("  {");
                            sw.WriteLine("    \"DIAMETER\"" + "                           " + D);
                            sw.WriteLine("    \"PLATE_THICKNESS\"" + "                           " + t);
                            sw.WriteLine("  }");
                            sw.WriteLine("}");
                            sw.WriteLine("");
                        }
                    else if (section_type == 'T')
                        {
                            double H = Convert.ToDouble(profileName.Split('*')[0].Substring(1));
                            double W = Convert.ToDouble(profileName.Split('*')[1]);
                            double F = Convert.ToDouble(profileName.Split('*')[2]);
                            double Y = Convert.ToDouble(profileName.Split('*')[3]);

                            sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                            sw.WriteLine("{");
                            sw.WriteLine("  TYPE = 10; SUB_TYPE = 10001; COORDINATE = 0.000;");
                            sw.WriteLine("  {");
                            sw.WriteLine("    \"HEIGHT\"" + "                           " + H);
                            sw.WriteLine("    \"WIDTH\"" + "                           " + W);
                            sw.WriteLine("    \"WEB_THICKNESS\"" + "                           " + F);
                            sw.WriteLine("    \"FLANGE_THICKNESS\"" + "                           " + Y);
                            sw.WriteLine("    \"ROUNDING_RADIUS_1\"" + "                           " + "0.000000000E+000");
                            sw.WriteLine("    \"ROUNDING_RADIUS_2\"" + "                           " + "0.000000000E+000");
                            sw.WriteLine("    \"FLANGE_SLOPE_RATIO\"" + "                           " + "0.000000000E+000");
                            sw.WriteLine("  }");
                            sw.WriteLine("}");
                            sw.WriteLine("");
                        }





            }
                

                sw.Close();
                fil.Close();

                #endregion





        }

        /// <summary>
        /// 获得数据库数据
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmd"></param>
        /// <returns></returns>
        public DataTable getData(string sql, SQLiteCommand cmd)
        {

            cmd.CommandText = sql;

            SQLiteDataAdapter adapter= new SQLiteDataAdapter(cmd);
            DataSet ds = new DataSet();
            adapter.Fill(ds);

            DataTable table = ds.Tables[0];

            return table;
        }

        /// <summary>
        /// 得到新的偏移点
        /// </summary>
        /// <param name="point"></param>
        /// <param name="VectorPoint"></param>
        /// <param name="DevString"></param>
        /// <returns></returns>
        public Tekla.Structures.Geometry3d.Point GetNewStartPoint(Tekla.Structures.Geometry3d.Point point, Tekla.Structures.Geometry3d.Point VectorPoint,double DevString)
        {


            Tekla.Structures.Geometry3d.Point pt = new Tekla.Structures.Geometry3d.Point();

            pt.X = (-VectorPoint.Y * DevString)/ Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.X;
            pt.Y = (VectorPoint.X * DevString ) / Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.Y;
            pt.Z = point.Z;

            return pt;
        }

        public Tekla.Structures.Geometry3d.Point GetNewStartPointSecond(Tekla.Structures.Geometry3d.Point point, Tekla.Structures.Geometry3d.Point VectorPoint, double DevString)
        {


            Tekla.Structures.Geometry3d.Point pt = new Tekla.Structures.Geometry3d.Point();

            pt.X = (-VectorPoint.X * DevString) / Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.X;
            pt.Y = (VectorPoint.Y * DevString) / Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.Y;
            pt.Z = point.Z;

            return pt;
        }

        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Output_Text_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
