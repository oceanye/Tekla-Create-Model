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
using Point = Tekla.Structures.Geometry3d.Point;
using Vector = Tekla.Structures.Geometry3d.Vector;
using System.Text.RegularExpressions;
using Tekla.Structures.Drawing;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula.Functions;
using Match = System.Text.RegularExpressions.Match;
using ModelObjectSelector = Tekla.Structures.Model.UI.ModelObjectSelector;
using static Tekla.Structures.Model.Part;


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

            Refresh_Material_List();

            //combo_color增加两个选项("根据截面区分", "根据梁柱区分")，默认选择 根据截面
            combo_color.Items.Add("截面区分");
            combo_color.Items.Add("梁柱区分");
            combo_color.SelectedIndex = 0;


        }


        public string DbPath = @"C:\ProgramData\Autodesk\Revit\Addins\2018";




        List<string> ProfileList = new List<string>();
        List<string> ProfileList_beam = new List<string>();
        List<string> ProfileList_column = new List<string>();

        List<string> PList = new List<string>();


        //  public string DbPath = @"F:\Tekla\RvtData_v0526_debug.db";



        private void button_Click(object sender, RoutedEventArgs e)
        {

            //清空Output_Text
            Output_Text.Clear();

            bool PIP_in_Centre = (bool)checkbox_PIP_CENTRE.IsChecked;


            Generate_Model();
            if (PIP_in_Centre)
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

                //ProfileList 新模型所需截面

                List<List<string>> StandardProfileList = Import_Standard_ProfileList("H");
                //import beam section

                for (int x = 0; x < profileBeamData.Rows.Count; x++)
                {
                    string Profile_temp = profileBeamData.Rows[x].ItemArray[1].ToString().Split(' ')[2].Replace('X', '*');

                    Profile_temp=Profile_Standardize(Profile_temp, StandardProfileList);

                    ProfileList_beam.Add(Profile_temp);
                }

                //建立一个dictionary，每一个profileList_beam的string对应一个int，从1开始，每次加1，跳过重复
                Dictionary<string, int> ProfileList_beam_color = new Dictionary<string, int>();
                int count = 1;
                foreach (string profile in ProfileList_beam)
                {
                    if (!ProfileList_beam_color.ContainsKey(profile))
                    {
                        ProfileList_beam_color.Add(profile, count);
                        count++;
                    }
                }



                // import columb section
                for (int x = 0; x < profileColumData.Rows.Count; x++)
                {
                    string Profile_temp = profileColumData.Rows[x].ItemArray[1].ToString().Split(' ')[2].Replace('X', '*');

                    Profile_temp = Profile_Standardize(Profile_temp, StandardProfileList);

                    ProfileList_column.Add(Profile_temp);
                }

                //建立一个dictionary，每一个profileList_beam的string对应一个int，从1开始，每次加1，跳过重复
                Dictionary<string, int> ProfileList_column_color = new Dictionary<string, int>();
                count = 1;
                foreach (string profile in ProfileList_column)
                {
                    if (!ProfileList_column_color.ContainsKey(profile))
                    {
                        ProfileList_column_color.Add(profile, count);
                        count++;
                    }
                }

                //join the beam and column profileList
                ProfileList.AddRange(ProfileList_beam);
                ProfileList.AddRange(ProfileList_column);


                List<LibraryProfileItem> profileL = new List<LibraryProfileItem>();
                List<MaterialItem> materialL = new List<MaterialItem>();

                CatalogHandler CatalogHandler = new CatalogHandler();



                ProfileItemEnumerator ProfileItemEnumerator = CatalogHandler.GetLibraryProfileItems();
                MaterialItemEnumerator MaterialItemEn = CatalogHandler.GetMaterialItems();


                //profileL 现有Tekla截面库

                while (ProfileItemEnumerator.MoveNext())
                {
                    LibraryProfileItem LibraryProfileItem = ProfileItemEnumerator.Current as LibraryProfileItem;

                    profileL.Add(LibraryProfileItem);

                }

                while (MaterialItemEn.MoveNext())
                {
                    MaterialItem materialItem = MaterialItemEn.Current as MaterialItem;

                    materialL.Add(materialItem);
                    //Material_List.Add(materialItem.MaterialName);

                }


                for (int i = 0; i < ProfileList.Count; i++)
                {
                    int Flag = 0;
                    for (int j = 0; j < profileL.Count; j++)
                    {
                        string Profile1 = ProfileList[i].Replace("*@PEC", "");
                       //if (ProfileList[i].Contains(profileL[j].ProfileName))
                       if (profileL[j].ProfileName.Contains(Profile1))
                        {
                            Flag++;
                        }
                    }


                    //PList 需要新建的截面
                    if (Flag == 0) //现有截面库不包含，需要新建导入
                    {
                        PList.Add(ProfileList[i]);
                    }

                }

                if (PList.Count > 0)
                {
                    Model myModel = new Model();
                    string path = myModel.GetInfo().ModelPath + "\\";
                    CreateNewProfileFile(path, PList);
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

                #region 生成新标高
                cmd.CommandText = "SELECT*FROM LevelTable";

                SQLiteDataAdapter adapterLevel = new SQLiteDataAdapter(cmd);
                DataSet dsLevel = new DataSet();
                adapterLevel.Fill(dsLevel);

                DataTable tableLevel = dsLevel.Tables[0];

                Model modelNewGrid = new Model();

                Tekla.Structures.Model.Grid grid = new Tekla.Structures.Model.Grid();

                string Z = "";

                string LabelZ = "+0 ";

                for (int i = 0; i < tableLevel.Rows.Count; i++)
                {
                    Z += tableLevel.Rows[i].ItemArray[2].ToString() + " ";

                    LabelZ += tableLevel.Rows[i].ItemArray[1].ToString() + " ";
                }

                #endregion

                #region 生成新轴网
                try
                {
                    cmd.CommandText = "SELECT*FROM GridTable";

                    SQLiteDataAdapter adapterGrid = new SQLiteDataAdapter(cmd);
                    DataSet dsGrid = new DataSet();
                    adapterLevel.Fill(dsGrid);

                    DataTable tableGrid = dsGrid.Tables[0];

                    List<string> LabelList = new List<string>();
                    List<string> gridPointList = new List<string>();

                    List<List<string>> gridInfo = new List<List<string>>();



                    //Console.WriteLine(LabelList);
                    for (int i = 0; i < tableGrid.Rows.Count; i++)
                    {
                        string column1Value = tableGrid.Rows[i][1].ToString();
                        string column2Value = tableGrid.Rows[i][2].ToString().Split(';')[0];
                        string column3Value = tableGrid.Rows[i][2].ToString().Split(';')[1];

                        List<string> row = new List<string> { column1Value, column2Value, column3Value };

                        gridInfo.Add(row);

                        //LabelList.Add( tableGrid.Rows[i].ItemArray[1].ToString() );

                        //gridPointList.Add(tableGrid.Rows[i].ItemArray[2].ToString());
                    }


                    List<List<string>> letterList = new List<List<string>>();
                    List<List<string>> numberList = new List<List<string>>();

                    foreach (List<string> sublist in gridInfo)
                    {

                        // 使用 LINQ 查询来检查元素是否包含数字

                        if (ParallelAxis(sublist[1], sublist[2]) == true) // 只保留平行于坐标轴的轴网 
                        {
                            if (sublist[0].Any(char.IsDigit))
                            {
                                numberList.Add(sublist);
                            }
                            else
                            {
                                letterList.Add(sublist);
                            }

                        }


                    }
                    // letter - Y
                    // number - X

                    List<double> dist_letterList = new List<double>();
                    List<double> dist_numberList = new List<double>();

                    List<string> label_letterList = new List<string>();
                    List<string> label_numberList = new List<string>();

                    label_letterList = letterList.Select(sublist => sublist.First()).ToList();
                    label_numberList = numberList.Select(sublist => sublist.First()).ToList();

                    dist_letterList.Add(ParsePointString(letterList[0][1]).Y);
                    dist_numberList.Add(ParsePointString(numberList[0][1]).X);

                    for (int i = 1; i < letterList.Count; i++)
                    {

                        double d1 = DistaceAxis(letterList[i - 1][1], letterList[i - 1][2], letterList[i][1]);
                        dist_letterList.Add(d1);


                    }

                    for (int i = 1; i < numberList.Count; i++)
                    {

                        double d1 = DistaceAxis(numberList[i - 1][1], numberList[i - 1][2], numberList[i][1]);
                        dist_numberList.Add(d1);

                    }

                    //List<string> letterList;
                    //List<string> numberList;
                    //List<double> List_dist_A;
                    //List<double> List_dist_1;

                    //ConvertGrid(LabelList, gridPointList, out letterList, out numberList, out List_dist_A, out List_dist_1);

                    grid.Name = "Grid";
                    grid.CoordinateX = string.Join(" ", dist_numberList);
                    grid.CoordinateY = string.Join(" ", dist_letterList);
                    grid.CoordinateZ = Z;
                    grid.LabelX = string.Join(" ", label_numberList);
                    grid.LabelY = string.Join(" ", label_letterList);
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
                }
                catch(Exception ex)
                {

                }
                

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
                    string ProfileBeam_org = section_name.Split(' ')[2];
                    string ProfileBeam = ProfileBeam_org.Replace('X', '*');//.Substring(1);

                    //在ProfileList_Beam中匹配包含 ProfileBeam的项，如果存在，将ProfileBeam赋值为ProfileList_Beam中的项，如果不存在，保留原值
                    ProfileBeam = Profile_Standardize(ProfileBeam, StandardProfileList);



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
                        if (profileL[x].ProfileName==(ProfileBeam))
                        {
                            beam.Profile.ProfileString = profileL[x].ProfileName;
                            break;
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
                        beam.Material.MaterialString = combox_mat_beam.SelectedItem.ToString();
                    }

                    // 根据combo_color的选择，建立switch case是"构件区分"，和case是"截面区分”
                    switch (combo_color.SelectionBoxItem.ToString())
                    {
                        case "梁柱区分":
                            beam.Class = "3";
                            break;
                        case "截面区分":
                            int color_index = 0;
                            try 
                            { color_index = ProfileList_beam_color[beam.Profile.ProfileString]; }
                            catch(Exception)
                            { System.Windows.Forms.MessageBox.Show("检查截面"+ ProfileBeam); }

                            beam.Class = color_index.ToString();
                            break;
                    }


                    //将int转换为string


                    if (comboBox.SelectionBoxItem.ToString() == "是")// 此段可以删除，中心线始终位于钢梁翼缘顶部
                    {
                        beam.StartPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[11]);
                        beam.EndPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[12]);

                        beam.StartPointOffset.Dz = Convert.ToDouble(tableBeam.Rows[i].ItemArray[15]);
                        beam.EndPointOffset.Dz = Convert.ToDouble(tableBeam.Rows[i].ItemArray[16]);
                    }
                    else
                    {
                        //Revit导出时，设置Y项对正为 LocationLine
                        //beam.StartPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[11]);
                        //beam.EndPointOffset.Dy = Convert.ToDouble(tableBeam.Rows[i].ItemArray[12]);

                        //beam.StartPoint.Z = beam.StartPoint.Z + Convert.ToDouble(tableBeam.Rows[i].ItemArray[15]);
                        //beam.EndPoint.Z = beam.EndPoint.Z + Convert.ToDouble(tableBeam.Rows[i].ItemArray[16]);
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
                            Console.WriteLine("Hole：" + (holetable.Rows[x].ItemArray[0].ToString()));

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
                                Chamfer chamfer = new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_ARC_POINT);

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

                            Console.WriteLine("CP:" + CP.Identifier);
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
                Output_Text.AppendText("生成模型 梁类型:" + ProfileList_beam_color.Count + "种,构件 " + T + "个 \n");

                Output_Text.AppendText("包含梁上开洞" + (Rec_hole + Cir_hole) + "个：方" + Rec_hole + "/圆" + Cir_hole + "个 \n");


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





                    Beam column = new Beam();

                    string startStr = table.Rows[i].ItemArray[1].ToString().Substring(1, table.Rows[i].ItemArray[1].ToString().Length - 2);

                    string endStr = table.Rows[i].ItemArray[2].ToString().Substring(1, table.Rows[i].ItemArray[2].ToString().Length - 2);

                    string ProfileColumn_org = table.Rows[i].ItemArray[3].ToString().Split(' ')[2];
                    string ProfileColumn = ProfileColumn_org.Replace('X', '*');//.Substring(1);

                    ProfileColumn = Profile_Standardize(ProfileColumn, StandardProfileList);

                    //List<List<string>> StandardProfileList = Import_Standard_ProfileList();
                    //string ProfilePrex = Profile_Check(ProfileColumn, StandardProfileList);

                    string MaterialColumn = table.Rows[i].ItemArray[11].ToString();

                    column.StartPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(startStr.Split(',')[0]), Convert.ToDouble(startStr.Split(',')[1]), Convert.ToDouble(startStr.Split(',')[2]));

                    column.EndPoint = new Tekla.Structures.Geometry3d.Point(Convert.ToDouble(endStr.Split(',')[0]), Convert.ToDouble(endStr.Split(',')[1]), Convert.ToDouble(endStr.Split(',')[2]));

                    column.Profile.ProfileString = "";

                    for (int x = 0; x < profileL.Count; x++)
                    {
                        if (profileL[x].ProfileName.Contains(ProfileColumn))
                        {
                            column.Profile.ProfileString = profileL[x].ProfileName;//ProfileColumn;
                        }
                    }

                    column.Position.Plane = Position.PlaneEnum.MIDDLE;

                    column.Position.Depth = Position.DepthEnum.MIDDLE;


                    column.Position.Rotation = Position.RotationEnum.TOP;


                    column.Position.RotationOffset = Convert.ToDouble(table.Rows[i].ItemArray[10].ToString()) - 90;


                    int flag = 0;
                    for (int x = 0; x < materialL.Count; x++)
                    {
                        if (materialL[x].MaterialName.Contains(MaterialColumn))
                        {
                            column.Material.MaterialString = MaterialColumn;
                            flag++;
                        }
                    }
                    if (flag == 0)
                    {
                        column.Material.MaterialString = combox_mat_column.SelectedItem.ToString();
                    }


                    // 根据combo_color的选择，建立switch case是"构件区分"，和case是"截面区分”
                    switch (combo_color.SelectionBoxItem.ToString())
                    {
                        case "梁柱区分":
                            column.Class = "3";
                            break;
                        case "截面区分":
                            column.Class = ProfileList_column_color[column.Profile.ProfileString].ToString();
                            break;
                    }

                    column.Name = "column";


                    if (column.Profile.ProfileString != "")
                    {
                        column.Insert();
                    }




                    modelCo.CommitChanges();
                }

                Output_Text.AppendText("生成模型 柱类型:" + ProfileList_column_color.Count + "种，构件" + T1 + "个\n");


                #endregion


                conn.Close();





            }







            //Close();




        }



        //建立一个函数phraseHSection(section_name),通过对H型截面的解析，返回截面的H,B,tw,tf。通过section_name的字段进行正则匹配，第一个数字是H，第二个数字是B，第三个数字是tw，第四个数字是tf，返回一个list
        public List<double> phraseHSection(string section_name)
        {
            List<double> sectionValues = new List<double>();

            Match match = Regex.Match(section_name, @"([A-Za-z]{1,2})(\d+)\D*(\d+)\D*(\d+)\D*(\d+)\D*");

            if (match.Success)
            {
                // 使用捕获组提取四个数字

                    //要求是int sectionValues.AddRange((match.Groups.Cast<Group>().Skip(2).Select(g => g.Value)));
                    sectionValues.AddRange((match.Groups.Cast<Group>().Skip(2).Select(g => Convert.ToDouble(g.Value))));

            }

            return sectionValues;
        }

        private void CreatePEC(Beam beam)
        {
            Model model = new Model();
            Beam beam_C1 = new Beam();

            //完成如下功能 H B tw tf = phraseHSection(beam.Profile.ProfileString)
            List<double> sectionValues = phraseHSection(beam.Profile.ProfileString);
            double H = sectionValues[0];
            double B = sectionValues[1];
            double tw = sectionValues[2];
            double tf = sectionValues[3];



            //在同样位置，加入一个截面是h=300,b=100的混凝土梁
            beam_C1.Profile.ProfileString = (H - 2 * tf).ToString() + "*" + (B / 2 - tw / 2).ToString();
            beam_C1.Material.MaterialString = "C30";
            beam_C1.Class = "1";
            beam_C1.StartPoint = beam.StartPoint;
            beam_C1.EndPoint = beam.EndPoint;

            beam_C1.StartPointOffset.Dy = (B / 4) + (tw / 4);
            beam_C1.EndPointOffset.Dy = (B / 4) + (tw / 4);
            beam_C1.StartPointOffset.Dz = -tf;
            beam_C1.EndPointOffset.Dz = -tf;
            beam_C1.Insert();

            Beam beam_C2 = new Beam();

            beam_C2.Profile.ProfileString = (H - 2 * tf).ToString() + "*" + (B / 2 - tw / 2).ToString();
            beam_C2.Material.MaterialString = "C30";
            beam_C2.Class = "1";
            beam_C2.StartPoint = beam.StartPoint;
            beam_C2.EndPoint = beam.EndPoint;

            beam_C2.StartPointOffset.Dy = -(B / 4 + tw / 4);
            beam_C2.EndPointOffset.Dy = -(B / 4 + tw / 4);
            beam_C2.StartPointOffset.Dz = -tf;
            beam_C2.EndPointOffset.Dz = -tf;
            beam_C2.Insert();


            // Create main assembly
            Assembly mainAssembly = new Assembly();
            mainAssembly.Name = "MainAssembly";
            mainAssembly.Insert();

            mainAssembly.Add(beam_C1);
            mainAssembly.SetMainPart(beam_C1);
            
            mainAssembly.Add(beam_C2);
            mainAssembly.Modify();

            
            //找到包含beam的assembly，并加入到mainAssembly
            //mainAssembly.Add(beam.GetAssembly());



            model.CommitChanges();
        }

        

        static void ConvertGrid(List<string> LabelList, List<string> gridPointList, out List<string> letterList, out List<string> numberList, out List<double> List_dist_A, out List<double> List_dist_1)
        {
            letterList = new List<string>();
            numberList = new List<string>();
            List_dist_A = new List<double>();
            List_dist_1 = new List<double>();

            foreach (string label in LabelList)
            {
                if (char.IsLetter(label[0]))
                {
                    letterList.Add(label);
                }
                else if (char.IsDigit(label[0]))
                {
                    numberList.Add(label);
                }
            }


            List<List<string>> group = null;
            foreach (string label in letterList)
            {
                group.Add(SelectCoordbyLabel(label, LabelList, gridPointList));
            }



            // List_dist_A.Add(distance);


            foreach (string label in numberList)
            {
                //List<string> group = GetGroup(label, LabelList, gridPointList);
                if (group.Count >= 2)
                {
                    //    double distance = CalculateDistance(group[0], group[1]);
                    //   List_dist_1.Add(distance);
                }
            }
        }

        static List<string> SelectCoordbyLabel(string label, List<string> LabelList, List<string> gridPointList)
        {
            List<String> Point_Pair = null;
            int index = LabelList.IndexOf(label);
            Point_Pair.Add(label);
            Point_Pair.AddRange(gridPointList[index].Split(';').ToList());
            return Point_Pair;
            //return LabelList.Where(point => point.Contains(label)).ToList();
        }

        static double CalculateDistance(string point1, string point2)
        {
            // 解析字符串并提取坐标
            Point p1 = ParsePointString(point1);
            Point p2 = ParsePointString(point2);

            // 计算距离
            Vector vector = new Vector(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z);
            return vector.GetLength();
        }



        public static String Profile_Standardize(string Profile_org, List<List<string>> StandardProfileList)
        {



            string Profile_new = "";

            string section_type = "";   //截面类型
            string section_size = ""; //截面尺寸
            string section_type_standard = ""; //查表后标准截面类型



            Match matches = Regex.Match(Profile_org, @"^([A-Za-z]+)(\d.*)$");

            List<string> profileValues = new List<string>();

            section_type=(matches.Groups[1].Value);
            section_size=(matches.Groups[2].Value);

            if (matches.Success)
            {
                section_type = (matches.Groups[1].Value);
                section_size = (matches.Groups[2].Value);

            }


            foreach (var item in StandardProfileList)
            {
                if (item[1] == section_size)
                {
                    section_type_standard = item[0];
                    break;
                }
            }
            if (section_type_standard == "")
            {
                section_type_standard = "BH";
            }

            if(section_type =="H")
            {
                Profile_new = section_type_standard + section_size;
            }
            else // 后续可拓展方形、圆形、T型等
            {
                Profile_new = Profile_org;
            }

            return Profile_new;

        }


        public static List<List<string>> Import_Standard_ProfileList(string shape_type)
        {
            List<List<string>> Standard_ProfileList = new List<List<string>>();

            string filePath = "";
            //读取文件StandardProfileList.txt
            
            if (shape_type == "H")
            {
                filePath = "D://Program Files//GBS_Software//CreateTeklaModel//H_Profile_Standard.txt";
            }
            
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                foreach( string line in lines)
                {
                    string[] values = line.Split(',');
                    Match matches = Regex.Match(line, @"^([A-Za-z]+)(\d.*)$");

                    List<string> profileValues = new List<string>();

                    profileValues.Add(matches.Groups[1].Value);
                    profileValues.Add(matches.Groups[2].Value);
                    
                    Standard_ProfileList.Add(profileValues);
                }


            }
            catch(Exception ex)
            {

            }
            return Standard_ProfileList;
        }

    static bool ParallelAxis(string point1, string point2)
        {
            Point p1 = ParsePointString(point1);
            Point p2 = ParsePointString(point2);

            // 计算距离
            Vector vector = new Vector(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z);

            bool isParallelToXAxis = Math.Abs(vector.Y) < 1e-3;
            bool isParallelToYAxis = Math.Abs(vector.X) < 1e-3;

            if (isParallelToXAxis || isParallelToYAxis)
            {
                return true;
            }
            else
            {
                return false;
            }



        }

        static double DistaceAxis(string point1, string point2, string point3)
        {
            Point p1 = ParsePointString(point1);
            Point p2 = ParsePointString(point2);
            Point p3 = ParsePointString(point3);

            Tekla.Structures.Geometry3d.Line L1 = new Tekla.Structures.Geometry3d.Line(p1, p2);
            double d1 = Tekla.Structures.Geometry3d.Distance.PointToLine(p3, L1);

            return d1;
        }

        static Point ParsePointString(string pointString)
        {
            //string[] coordinates = pointString.Split(';');
            //if (coordinates.Length == 2)
            {
                string[] coord1 = pointString.Trim('(', ')').Split(',');
                //string[] coord2 = pointString[1].Trim('(', ')').Split(',');
                if (coord1.Length == 3)
                {
                    double x1 = double.Parse(coord1[0]);
                    double y1 = double.Parse(coord1[1]);
                    double z1 = double.Parse(coord1[2]);

                    return new Point(x1, y1, z1);
                }
            }
            throw new ArgumentException("无法解析的坐标字符串: " + pointString);
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



    public void CreateNewProfileFile(string filePath, List<string> pList)
        {
            if (Directory.Exists(filePath) == false)
            {

                Directory.CreateDirectory(filePath);

            }

            //List<List<string>> StandardProfileList = Import_Standard_ProfileList("H");

            #region 


            FileStream fil = new FileStream(filePath + "截面文件.lis", FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fil);
            sw.WriteLine("PROFILE DATABASE EXPORT VERSION = 3");
            sw.WriteLine("");
            for (int i = 0; i < pList.Count; i++)
            {
                string profileName = pList[i].Replace("X", "*");

                string section_type = "";//Profile_Check(profileName, StandardProfileList);

                string section_size = "";

                Match matches = Regex.Match(profileName, @"^([A-Za-z]+)(\d.*)$");

                List<string> profileValues = new List<string>();
                if (matches.Success)
                {
                    section_type = matches.Groups[1].Value;
                    section_size = matches.Groups[2].Value;
                }



                if (section_type.Contains( 'H'))
                {
                    double H = Convert.ToDouble(section_size.Split('*')[0]);
                    double W = Convert.ToDouble(section_size.Split('*')[1]);
                    double F = Convert.ToDouble(section_size.Split('*')[2]);
                    double Y = Convert.ToDouble(section_size.Split('*')[3]);

                    double Area=Convert.ToDouble(H*W-(H-2*Y)*(W-F));
                    double Weight_Length= Convert.ToDouble(Area*7850*1e-6);

                    string profileType = "1001";

                    if (section_type.Contains("BH"))
                    {
                        profileType = "1002";
                    }

                    sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                    sw.WriteLine("{");
                    sw.WriteLine("  TYPE = 1; SUB_TYPE = "+profileType+"; COORDINATE = 0.000;");
                    sw.WriteLine("  {");
                    sw.WriteLine("    \"HEIGHT\"" + "                           " + H);
                    sw.WriteLine("    \"WIDTH\"" + "                           " + W);
                    sw.WriteLine("    \"WEB_THICKNESS\"" + "                           " + F);
                    sw.WriteLine("    \"FLANGE_THICKNESS\"" + "                           " + Y);
                    sw.WriteLine("    \"ROUNDING_RADIUS_1\"" + "                           " + "0.000000000E+000");
                    sw.WriteLine("    \"ROUNDING_RADIUS_2\"" + "                           " + "0.000000000E+000");
                    sw.WriteLine("    \"FLANGE_SLOPE_RATIO\"" + "                           " + "0.000000000E+000");
                    sw.WriteLine("    \"CROSS_SECTION_AREA\"" + "                           " + Area);
                    sw.WriteLine("    \"WEIGHT_PER_UNIT_LENGTH\"" + "                           " + Weight_Length);
                    sw.WriteLine("  }");
                    sw.WriteLine("}");
                    sw.WriteLine("");
                }
                else if (section_type == "B")
                {
                    double H = Convert.ToDouble(section_size.Split('*')[0]);
                    double W = Convert.ToDouble(section_size.Split('*')[1]);
                    double F = Convert.ToDouble(section_size.Split('*')[2]);
                    //double Y = Convert.ToDouble(profileName.Split('*')[3]);

                    double Area = Convert.ToDouble(H * W - (H - 2 * F) * (W - 2*F));
                    double Weight_Length = Convert.ToDouble(Area * 7850 * 1e-6);


                    sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                    sw.WriteLine("{");
                    sw.WriteLine("  TYPE = 8; SUB_TYPE = 8002; COORDINATE = 0.000;");
                    sw.WriteLine("  {");
                    sw.WriteLine("    \"HEIGHT\"" + "                           " + H);
                    sw.WriteLine("    \"WIDTH\"" + "                           " + W);
                    sw.WriteLine("    \"PLATE_THICKNESS\"" + "                           " + F);
                    sw.WriteLine("    \"ROUNDING_RADIUS\"" + "                           " + "1.000000000E+001");
                    sw.WriteLine("    \"CROSS_SECTION_AREA\"" + "                           " + Area);
                    sw.WriteLine("    \"WEIGHT_PER_UNIT_LENGTH\"" + "                           " + Weight_Length);
                    sw.WriteLine("  }");
                    sw.WriteLine("}");
                    sw.WriteLine("");
                }
                else if (section_type == "P")
                {
                    double D = Convert.ToDouble(section_size.Split('*')[0]);
                    double t = Convert.ToDouble(section_size.Split('*')[1]);

                    double Area = Convert.ToDouble((D * D - (D - 2 * t) * (D - 2 * t))/4*3.14);
                    double Weight_Length = Convert.ToDouble(Area * 7850 * 1e-6);



                    sw.WriteLine("PROFILE_NAME = " + "\"" + profileName + "\"" + ";");
                    sw.WriteLine("{");
                    sw.WriteLine("  TYPE = 7; SUB_TYPE = 7001; COORDINATE = 0.000;");
                    sw.WriteLine("  {");
                    sw.WriteLine("    \"DIAMETER\"" + "                           " + D);
                    sw.WriteLine("    \"PLATE_THICKNESS\"" + "                           " + t);
                    sw.WriteLine("    \"CROSS_SECTION_AREA\"" + "                           " + Area);
                    sw.WriteLine("    \"WEIGHT_PER_UNIT_LENGTH\"" + "                           " + Weight_Length);
                    sw.WriteLine("  }");
                    sw.WriteLine("}");
                    sw.WriteLine("");
                }
                else if (section_type == "T")
                {
                    double H = Convert.ToDouble(section_size.Split('*')[0]);
                    double W = Convert.ToDouble(section_size.Split('*')[1]);
                    double F = Convert.ToDouble(section_size.Split('*')[2]);
                    double Y = Convert.ToDouble(section_size.Split('*')[3]);


                    double Area = Convert.ToDouble((H * W - (H - 2 * F) * (W - 1 * Y)));
                    double Weight_Length = Convert.ToDouble(Area * 7850 * 1e-6);



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
                    sw.WriteLine("    \"CROSS_SECTION_AREA\"" + "                           " + Area);
                    sw.WriteLine("    \"WEIGHT_PER_UNIT_LENGTH\"" + "                           " + Weight_Length);
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

            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
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
        public Tekla.Structures.Geometry3d.Point GetNewStartPoint(Tekla.Structures.Geometry3d.Point point, Tekla.Structures.Geometry3d.Point VectorPoint, double DevString)
        {


            Tekla.Structures.Geometry3d.Point pt = new Tekla.Structures.Geometry3d.Point();

            pt.X = (-VectorPoint.Y * DevString) / Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.X;
            pt.Y = (VectorPoint.X * DevString) / Math.Sqrt(VectorPoint.X * VectorPoint.X + VectorPoint.Y * VectorPoint.Y) + point.Y;
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

        private void ComboBox_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {

        }

        public void combox_mat_beam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            

        }

        public void CopyMaterial()
        {
            try
            {
                CatalogHandler catalogHandler = new CatalogHandler();
                //catalogHandler.GetConnection();

                MaterialItemEnumerator MaterialItemEn = catalogHandler.GetMaterialItems();

                //MaterialItem newMaterial = new MaterialItem();



                //Refresh_Material_List();
                //Material xxx = catalogHandler.GetMaterial("Q345");
                 while (MaterialItemEn.MoveNext())
                {
                MaterialItem materialItem = MaterialItemEn.Current as MaterialItem;
                //materialL.Add(materialItem);
                string mat_temp = materialItem.MaterialName;
                     if (materialItem.MaterialName.Contains("Q345"))
                    {

                            MaterialItem newMaterial = materialItem.Copy();

                            newMaterial.MaterialName = MaterialTextBox.Text;
                            newMaterial.PlateDensity = 7850;
                            newMaterial.ProfileDensity = 7850;
                            newMaterial.Insert();

                            //newMaterial.MaterialName = "111";
                            //newMaterial.Insert();

                            //bool t = newMaterial.Modify();

                            //MaterialItem newMaterial=new MaterialItem();
                            //newMaterial.Insert();
                            //newMaterial.MaterialName = "Q355";
                            //newMaterial.Copy();

                            break;
                     }

                 }

            }
            catch
            {

            }
            Refresh_Material_List();
        }

        private void Refresh_Material_List()
        {
            CatalogHandler CatalogHandler = new CatalogHandler();



            //ProfileItemEnumerator ProfileItemEnumerator = CatalogHandler.GetLibraryProfileItems();
            MaterialItemEnumerator MaterialItemEn = CatalogHandler.GetMaterialItems();

            List<string> Material_List = new List<string>();



            while (MaterialItemEn.MoveNext())
            {
                MaterialItem materialItem = MaterialItemEn.Current as MaterialItem;

                //materialL.Add(materialItem);
                string mat_temp = materialItem.MaterialName;
                if (mat_temp.StartsWith("Q"))
                {
                    Material_List.Add(materialItem.MaterialName);
                }



            }

            string Default_Mat = "Q235";

            if (Material_List.Contains("Q355"))

            {
                Default_Mat = "Q355";
            }


            combox_mat_column.ItemsSource = Material_List;
            combox_mat_beam.ItemsSource = Material_List;

            combox_mat_column.SelectedItem = Default_Mat;
            combox_mat_beam.SelectedItem = Default_Mat;


        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            CopyMaterial();
        }

        private void SelectPEC_Click(object sender, RoutedEventArgs e)
        {
            //Model model = new Model();

            // 选择用户框选的梁对象
            // 创建一个 ModelObjectSelector 对象来获取当前选中的对象
            ModelObjectSelector selector = new ModelObjectSelector();

            // 获取当前选中的对象
            ModelObjectEnumerator selectedObjects = selector.GetSelectedObjects();

            // 历遍每个选中的对象
            while (selectedObjects.MoveNext())
            {
                Beam selectedBeam = selectedObjects.Current as Beam;
                if (selectedBeam != null)
                {
                    CreatePEC(selectedBeam);
                }
            }
        }
    }
}
