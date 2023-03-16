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
using Tekla.Structures;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;


using Tekla.Structures.Catalogs;
using System.Diagnostics;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SQLite;
using System.Data;

namespace TestTekla
{
    /// <summary>
    /// GetModuelPosition.xaml 的交互逻辑
    /// </summary>
    public partial class GetModuelPosition : Window
    {
        public GetModuelPosition()
        {
            InitializeComponent();
        }

        //  public string filePath = @"F:\Tekla\测试\";

        public string filePath;

        public string SiglefilePath ="";

        //  public string PyfilePath = @"Y:\数字化课题\数据库\py库调用\套料优化ortools\";
        public string PyfilePath = @"D:\Program Files\GBS_Software\CreateTeklaCNC";

        private void btn_Click(object sender, RoutedEventArgs e)
        {




            if (PyfilePath != "")
            {


                Model myModel = new Model();

                filePath = myModel.GetInfo().ModelPath + @"\合并文件\";
              


                ArrayList objectToSelectBeam = new ArrayList();

                List<Fitting> fittingList = new List<Fitting>();


                ModelObjectEnumerator selectObjects = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();

                //  ModelObjectEnumerator myEnumBeam = myModel.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.BEAM);
                ModelObjectEnumerator myEnumFitting = myModel.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.FITTING);


                if (selectObjects.GetSize() > 0)
                {
                    while (selectObjects.MoveNext())
                    {
                        objectToSelectBeam.Add(selectObjects.Current as Beam);
                    }
                }

                else
                {
                    System.Windows.Forms.MessageBox.Show("请选择构件！");

                    myModel.CommitChanges();

                    Close();
                    return;
                }

                //while (myEnumBeam.MoveNext())
                //{
                //    objectToSelectBeam.Add(myEnumBeam.Current as Beam);
                //}
                while (myEnumFitting.MoveNext())
                {
                    fittingList.Add(myEnumFitting.Current as Fitting);
                }




                //Tekla.Structures.Model.UI.ModelObjectSelector ms = new Tekla.Structures.Model.UI.ModelObjectSelector();
                //ms.Select(objectToSelect);

                List<List<Beam>> beamAllList = new List<List<Beam>>();

                List<List<Beam>> beamAllListSecond = new List<List<Beam>>();

                foreach (var item in objectToSelectBeam)
                {
                    #region 


                    List<Beam> beamList = new List<Beam>();
                    Beam beam = item as Beam;


                    if (beam != null)
                    {
                        double length = 0;


                        beam.GetReportProperty("LENGTH", ref length);
                        if (length <= 12000)
                        {

                            if (beam.Profile.ProfileString.Contains("H"))
                            {


                                foreach (var itemSecond in objectToSelectBeam)
                                {
                                    Beam beam1 = itemSecond as Beam;

                                    if (beam1 != null)
                                    {
                                        if (beam.Profile.ProfileString == beam1.Profile.ProfileString)
                                        {
                                            beamList.Add(beam1);
                                        }
                                    }


                                }
                            }

                            int T = 0;
                            for (int i = 0; i < beamAllList.Count; i++)
                            {

                                for (int j = 0; j < beamAllList[i].Count; j++)
                                {
                                    if (beamList.Contains(beamAllList[i][j]))
                                    {
                                        T++;
                                    }
                                }

                            }
                            if (T == 0)
                            {
                                if (beamList.Count != 0)
                                {
                                    beamAllList.Add(beamList);
                                }

                            }
                        }


                    }

                    #endregion

                }

                if (beamAllList.Count >= 1)
                {


                    AddDataToSqlite(beamAllList);//存入数据库
                                                 //  CreatePyText(beamAllList);//py文件

                    RunPythonScript();//调用py脚本




                    //  List<List<string>> GroupIdList = ReadTxt(PyfilePath);

                    string path = myModel.GetInfo().ModelPath + @"\EXCEL文件\";

                    List<List<string>> GroupIdList =  ReadDb(PyfilePath,path);


                    beamAllListSecond = GetGroupBeam(GroupIdList, objectToSelectBeam);

                 


                     CreateAllText(beamAllListSecond, fittingList);//分组文件


                   

                    // WriteExcel(beamAllListSecond, path);
                   

                    //  CreateSigleText(objectToSelectBeam);//单个文件




                    myModel.CommitChanges();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请选择构件种类大于1");
                    Close();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择文件夹");
            }

            Close();



        }



        public void CreateTxt(string fileName, string MateiName, string ProfileName, double Length, ArrayList objectToSelect, Tekla.Structures.Geometry3d.Point StartPoint)
        {

            DirectoryInfo dirs = new DirectoryInfo(filePath);
            DirectoryInfo[] dir = dirs.GetDirectories();
            FileInfo[] file = dirs.GetFiles();

            double radio = 0; ;

            List<List<Tekla.Structures.Geometry3d.Point>> pointL = new List<List<Tekla.Structures.Geometry3d.Point>>();

            if (Directory.Exists(filePath) == false)
            {

                Directory.CreateDirectory(filePath);

            }

            foreach (var bolitItem in objectToSelect)
            {

                List<Tekla.Structures.Geometry3d.Point> pointList = new List<Tekla.Structures.Geometry3d.Point>();
                List<Tekla.Structures.Geometry3d.Point> pointList1 = new List<Tekla.Structures.Geometry3d.Point>();
                List<Tekla.Structures.Geometry3d.Point> pointList2 = new List<Tekla.Structures.Geometry3d.Point>();
                BoltArray bolitArray = bolitItem as BoltArray;

                radio = bolitArray.BoltSize;

                ArrayList pointArrayList = bolitArray.BoltPositions;

                foreach (var pt in pointArrayList)
                {
                    Tekla.Structures.Geometry3d.Point point = pt as Tekla.Structures.Geometry3d.Point;
                    pointList.Add(point);

                }
                pointList = pointList.OrderBy(m => m.Y).ToList();
                for (int i = 0; i < pointList.Count; i++)
                {
                    if (i < pointList.Count / 2)
                    {
                        pointList1.Add(pointList[i]);
                    }
                    else
                    {
                        pointList2.Add(pointList[i]);
                    }

                }

                pointL.Add(pointList1);
                pointL.Add(pointList2);



            }



            string ProfileStr = Regex.Replace(ProfileName, @"^[A-Za-z]+", string.Empty);
            if (!File.Exists(filePath + fileName + ".nc1"))
            {
                FileStream fil = new FileStream(filePath + fileName + ".nc1", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);


                #region
                sw.WriteLine("ST");
                sw.WriteLine("**" + fileName + ".nc1");
                sw.WriteLine("  1");
                sw.WriteLine("  1");
                sw.WriteLine("  H1");
                sw.WriteLine("  H1");

                sw.WriteLine("  " + MateiName);
                sw.WriteLine("  3");
                sw.WriteLine("  " + ProfileName);
                sw.WriteLine("  I");
                sw.WriteLine("    " + Length);
                sw.WriteLine("     " + ProfileStr.Split('*')[0]);
                sw.WriteLine("      " + ProfileStr.Split('*')[1]);
                sw.WriteLine("       " + ProfileStr.Split('*')[3]);
                sw.WriteLine("       " + ProfileStr.Split('*')[2]);
                sw.WriteLine("       " + ProfileStr.Split('*')[3]);

                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine();
                sw.WriteLine();
                sw.WriteLine();
                sw.WriteLine();

                sw.WriteLine("AK");
                sw.WriteLine("  " + "v" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[0] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + ProfileStr.Split('*')[0] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");

                sw.WriteLine("AK");
                sw.WriteLine("  " + "o" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + "0.00" + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");


                sw.WriteLine("AK");
                sw.WriteLine("  " + "u" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");

                sw.WriteLine("BO");

                for (int i = 0; i < pointL.Count; i++)
                {
                    for (int j = 0; j < pointL[i].Count; j++)
                    {

                        sw.WriteLine("  " + "v" + "     " + Math.Round(Math.Abs(pointL[i][j].Y - StartPoint.Y), 2) + "o" + "    " + Math.Round(Math.Abs(pointL[i][j].Z - StartPoint.Z), 2) + "    " + Math.Round(radio, 2));

                    }
                }


                #endregion
                sw.Close();
                fil.Close();

            }
            else
            {
                FileStream fil = new FileStream(filePath + fileName + ".nc1", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);


                #region
                sw.WriteLine("ST");
                sw.WriteLine("**" + fileName + ".nc1");
                sw.WriteLine("  1");
                sw.WriteLine("  1");
                sw.WriteLine("  H1");
                sw.WriteLine("  H1");

                sw.WriteLine("  " + MateiName);
                sw.WriteLine("  3");
                sw.WriteLine("  " + ProfileName);
                sw.WriteLine("  I");
                sw.WriteLine("    " + Length);
                sw.WriteLine("     " + ProfileStr.Split('*')[0]);
                sw.WriteLine("      " + ProfileStr.Split('*')[1]);
                sw.WriteLine("       " + ProfileStr.Split('*')[3]);
                sw.WriteLine("       " + ProfileStr.Split('*')[2]);
                sw.WriteLine("       " + ProfileStr.Split('*')[3]);

                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine("       " + 0);
                sw.WriteLine();
                sw.WriteLine();
                sw.WriteLine();
                sw.WriteLine();

                sw.WriteLine("AK");
                sw.WriteLine("  " + "v" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[0] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + ProfileStr.Split('*')[0] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");

                sw.WriteLine("AK");
                sw.WriteLine("  " + "o" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + "0.00" + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");


                sw.WriteLine("AK");
                sw.WriteLine("  " + "u" + "       " + "0.00s" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("       " + Length + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + ProfileStr.Split('*')[1] + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");
                sw.WriteLine("         " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00" + "      " + "0.00");

                sw.WriteLine("BO");

                for (int i = 0; i < pointL.Count; i++)
                {
                    for (int j = 0; j < pointL[i].Count; j++)
                    {

                        sw.WriteLine("  " + "v" + "     " + Math.Round(Math.Abs(pointL[i][j].Y - StartPoint.Y), 2) + "o" + "    " + Math.Round(Math.Abs(pointL[i][j].Z - StartPoint.Z), 2) + "    " + Math.Round(radio, 2));

                    }
                }


                #endregion
                sw.Close();
                fil.Close();
            }


        }

        public void CreateText(string fileName, string ProfileName, double Length, ArrayList objectToSelect, Tekla.Structures.Geometry3d.Point StartPoint, int F, LibraryProfileItem currentProfileItem)
        {

            DirectoryInfo dirs = new DirectoryInfo(SiglefilePath);
            DirectoryInfo[] dir = dirs.GetDirectories();
            FileInfo[] file = dirs.GetFiles();

            double radio = 0; ;

            List<List<Tekla.Structures.Geometry3d.Point>> pointL = new List<List<Tekla.Structures.Geometry3d.Point>>();

            if (Directory.Exists(SiglefilePath) == false)
            {

                Directory.CreateDirectory(SiglefilePath);

            }

            foreach (var bolitItem in objectToSelect)
            {

                List<Tekla.Structures.Geometry3d.Point> pointList = new List<Tekla.Structures.Geometry3d.Point>();
                List<Tekla.Structures.Geometry3d.Point> pointList1 = new List<Tekla.Structures.Geometry3d.Point>();
                List<Tekla.Structures.Geometry3d.Point> pointList2 = new List<Tekla.Structures.Geometry3d.Point>();
                BoltArray bolitArray = bolitItem as BoltArray;

                radio = bolitArray.BoltSize;

                ArrayList pointArrayList = bolitArray.BoltPositions;

                foreach (var pt in pointArrayList)
                {
                    Tekla.Structures.Geometry3d.Point point = pt as Tekla.Structures.Geometry3d.Point;
                    pointList.Add(point);

                }
                pointList = pointList.OrderBy(m => Distance.PointToPoint(m, StartPoint)).ToList();
                for (int i = 0; i < pointList.Count; i++)
                {
                    if (i < pointList.Count / 2)
                    {
                        pointList1.Add(pointList[i]);
                    }
                    else
                    {
                        pointList2.Add(pointList[i]);
                    }

                }

                pointL.Add(pointList1);
                pointL.Add(pointList2);



            }



            // string ProfileStr = Regex.Replace(ProfileName, @"^[A-Za-z]+", string.Empty);
            if (!File.Exists(SiglefilePath + ProfileName.Substring(0, 1) + F + ".nc1"))
            {
                FileStream fil = new FileStream(SiglefilePath + ProfileName.Substring(0, 1) + F + ".nc1", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);


                #region
                double h = 0;
                double b = 0;
                double s = 0;
                double t = 0;
                foreach (var item in currentProfileItem.aProfileItemParameters)
                {
                    ProfileItemParameter par = item as ProfileItemParameter;
                    switch (par.Symbol)
                    {
                        case "h":
                            h = par.Value;
                            break;
                        case "b":
                            b = par.Value;
                            break;
                        case "s":
                            s = par.Value;
                            break;
                        case "t":
                            t = par.Value;
                            break;
                    }


                }


                sw.WriteLine("WORK SIZE :" + "L=" + Length * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + " TYPE=H");
                sw.WriteLine("");
                int flag = 0;
                for (int i = 0; i < pointL.Count; i++)
                {
                    for (int j = 0; j < pointL[i].Count; j++)
                    {
                        flag++;

                        sw.WriteLine("[" + flag + "]" + "X:" + Math.Round(Distance.PointToPoint(new Tekla.Structures.Geometry3d.Point(pointL[i][j].X, pointL[i][j].Y, StartPoint.Z), StartPoint), 2) + " Y:********* V:  " + Math.Round(Math.Abs(pointL[i][j].Z - StartPoint.Z), 2) + " B:********* T1:***** T2:D" + Math.Round(radio, 2) + "  T3:***** PAT:");

                    }
                }


                #endregion
                sw.Close();
                fil.Close();

            }
            else
            {
                FileStream fil = new FileStream(SiglefilePath + ProfileName.Substring(0, 1) + F + ".nc1", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);


                #region

                double h = 0;
                double b = 0;
                double s = 0;
                double t = 0;
                foreach (var item in currentProfileItem.aProfileItemParameters)
                {
                    ProfileItemParameter par = item as ProfileItemParameter;
                    switch (par.Symbol)
                    {
                        case "h":
                            h = par.Value;
                            break;
                        case "b":
                            b = par.Value;
                            break;
                        case "s":
                            s = par.Value;
                            break;
                        case "t":
                            t = par.Value;
                            break;
                    }
                }

                sw.WriteLine("WORK SIZE :" + "L=" + Length * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + " TYPE=H");
                sw.WriteLine("");
                int flag = 0;
                for (int i = 0; i < pointL.Count; i++)
                {
                    for (int j = 0; j < pointL[i].Count; j++)
                    {
                        flag++;
                        sw.WriteLine("[" + flag + "]" + "X:" + Math.Round(Distance.PointToPoint(new Tekla.Structures.Geometry3d.Point(pointL[i][j].X, pointL[i][j].Y, StartPoint.Z), StartPoint), 2) + " Y:********* V:  " + Math.Round(Math.Abs(pointL[i][j].Z - StartPoint.Z), 2) + " B:********* T1:***** T2:D" + Math.Round(radio, 2) + "  T3:***** PAT:");

                    }
                }


                #endregion
                sw.Close();
                fil.Close();
            }


        }


        public LibraryProfileItem GetProfile(string profileName)
        {
            LibraryProfileItem correntProfileItem = null;
            List<LibraryProfileItem> profileL = new List<LibraryProfileItem>();
            CatalogHandler CatalogHandler = new CatalogHandler();
            ProfileItemEnumerator ProfileItemEnumerator = CatalogHandler.GetLibraryProfileItems();


            while (ProfileItemEnumerator.MoveNext())
            {
                LibraryProfileItem LibraryProfileItem = ProfileItemEnumerator.Current as LibraryProfileItem;

                profileL.Add(LibraryProfileItem);

            }


            for (int j = 0; j < profileL.Count; j++)
            {
                if (profileL[j].ProfileName.Contains(profileName))
                {
                    correntProfileItem = profileL[j];
                }
            }

            return correntProfileItem;

        }


        public void CreateAllText(List<List<Beam>> beamAllList, List<Fitting> fittingList)
        {

            if (Directory.Exists(filePath) == false)
            {

                Directory.CreateDirectory(filePath);

            }


            MonitorCheck form = new MonitorCheck();
            form.Show();

            int T = 0;
            for (int i = 0; i < beamAllList.Count; i++)
            {
                string SumString = null;
                double SumLength = 0;

               

                for (int j = 0; j < beamAllList[i].Count; j++)
                {

                    double length = 0;
                    beamAllList[i][j].GetReportProperty("LENGTH", ref length);

                    if(j!= beamAllList[i].Count-1)
                    {
                        SumString += beamAllList[i][j].Identifier + "-";
                    }
                    else
                    {
                        SumString += beamAllList[i][j].Identifier;
                    }
                    
                    SumLength += length;
                   
                }

                LibraryProfileItem currentProfileItem = GetProfile(beamAllList[i][0].Profile.ProfileString);


                if (currentProfileItem != null)
                {


                    if (!File.Exists(filePath + "NCX_" + SumString + ".txt"))
                    {
                        T++;
                        form.SetTextMesssage(T, beamAllList.Count);//进度条函数

                        #region 


                        FileStream fil = new FileStream(filePath + "NCX_" + SumString + ".txt", FileMode.Create, FileAccess.Write);
                        StreamWriter sw = new StreamWriter(fil);

                        int Tag = 0;

                        #region 文件头
                        double h = 0;
                        double b = 0;
                        double s = 0;
                        double t = 0;
                        foreach (var item in currentProfileItem.aProfileItemParameters)
                        {
                            ProfileItemParameter par = item as ProfileItemParameter;
                            switch (par.Symbol)
                            {
                                case "h":
                                    h = par.Value;
                                    break;
                                case "b":
                                    b = par.Value;
                                    break;
                                case "s":
                                    s = par.Value;
                                    break;
                                case "t":
                                    t = par.Value;
                                    break;
                            }


                        }


                        sw.WriteLine("WORK SIZE :" + "L=" + Math.Round(SumLength, 2) * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + " TYPE=H" + " FNE:NC1");


                        #endregion
                        int flag = 0;

                        for (int j = 0; j < beamAllList[i].Count; j++)
                        {
                            double L = 0;
                            double CurrentL = 0;
                            if (j != 0)
                            {
                               // L = Distance.PointToPoint(beamAllList[i][j - 1].StartPoint, beamAllList[i][j - 1].EndPoint);

                                

                                for(int m=0;m< j;m++)
                                {
                                    double le = 0;
                                    beamAllList[i][j - (m+1)].GetReportProperty("LENGTH", ref le);

                                    L += le;


                                }
                               
                            }


                            for (int n = 0; n < j+1; n++)
                            {

                                double cL = 0;
                                beamAllList[i][n].GetReportProperty("LENGTH", ref cL);
                                CurrentL += cL;

                            }
                          
                           
                           // CurrentL = Distance.PointToPoint(beamAllList[i][j].StartPoint, beamAllList[i][j].EndPoint);
                            #region

                            ArrayList objectToSelect = new ArrayList();
                            ModelObjectEnumerator myEnum = beamAllList[i][j].GetBolts();

                            while (myEnum.MoveNext())
                            {
                                objectToSelect.Add(myEnum.Current as BoltArray);
                            }


                            if (objectToSelect.Count != 0)
                            {

                                flag= GetBoltPointList(objectToSelect, beamAllList[i][j].StartPoint, flag,j,sw, beamAllList,i, CurrentL, h,L, fittingList);

                            }
                            else
                            {

                                Tag++;
                                sw.Dispose();
                                fil.Dispose();
                                break;

                            }


                            #endregion

                        }

                        //if (Tag == 0)
                        //{
                        //    sw.Close();
                        //    fil.Close();
                        //    break;
                        //}

                           sw.Close();
                           fil.Close();

                        #endregion

                    }
                    else
                    {
                        T++;
                        form.SetTextMesssage(T, beamAllList.Count);//进度条函数
                        #region
                        FileStream fil = new FileStream(filePath + "NCX_" + SumString + ".txt", FileMode.Create, FileAccess.Write);
                        StreamWriter sw = new StreamWriter(fil);

                        int Tag = 0;

                        #region 文件头
                        double h = 0;
                        double b = 0;
                        double s = 0;
                        double t = 0;
                        foreach (var item in currentProfileItem.aProfileItemParameters)
                        {
                            ProfileItemParameter par = item as ProfileItemParameter;
                            switch (par.Symbol)
                            {
                                case "h":
                                    h = par.Value;
                                    break;
                                case "b":
                                    b = par.Value;
                                    break;
                                case "s":
                                    s = par.Value;
                                    break;
                                case "t":
                                    t = par.Value;
                                    break;
                            }


                        }


                        sw.WriteLine("WORK SIZE :" + "L=" + Math.Round(SumLength, 2) * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + "TYPE=H" + " FNE:NC1");



                        #endregion
                        int flag = 0;

                        for (int j = 0; j < beamAllList[i].Count; j++)
                        {
                            double L = 0;
                            double CurrentL = 0;
                            if (j != 0)
                            {
                                // L = Distance.PointToPoint(beamAllList[i][j - 1].StartPoint, beamAllList[i][j - 1].EndPoint);

                                for (int m = 0; m < j; m++)
                                {
                                    double le = 0;
                                    beamAllList[i][j - (m + 1)].GetReportProperty("LENGTH", ref le);

                                    L += le;


                                }
                            }


                            beamAllList[i][j].GetReportProperty("LENGTH", ref CurrentL);
                            #region

                            ArrayList objectToSelect = new ArrayList();
                            ModelObjectEnumerator myEnum = beamAllList[i][j].GetBolts();

                            while (myEnum.MoveNext())
                            {
                                objectToSelect.Add(myEnum.Current as BoltArray);
                            }


                            if (objectToSelect.Count != 0)
                            {

                                flag= GetBoltPointList(objectToSelect, beamAllList[i][j].StartPoint, flag, j, sw, beamAllList, i, CurrentL, h, L, fittingList);

                              

                             

                            }
                            else
                            {

                                Tag++;
                                sw.Dispose();
                                fil.Dispose();
                                break;

                            }


                            #endregion

                        }


                        //if (Tag == 0)
                        //{
                        //    sw.Close();
                        //    fil.Close();
                        //}

                        sw.Close();
                        fil.Close();
                        #endregion
                    }
                }
            }




        }


        public void CreateSigleText(ArrayList objectToSelectBeam)
        {
            int flag = 0;
            foreach (var item in objectToSelectBeam)
            {

                #region 

                ArrayList objectToSelect = new ArrayList();
                flag++;
                Beam BeamArray = item as Beam;
                ModelObjectEnumerator myEnum = BeamArray.GetBolts();



                while (myEnum.MoveNext())
                {
                    objectToSelect.Add(myEnum.Current as BoltArray);
                }



                string fileName = BeamArray.AssemblyNumber.Prefix + BeamArray.AssemblyNumber.StartNumber;
                string MateiName = BeamArray.Material.MaterialString;
                string ProfileName = BeamArray.Profile.ProfileString;

                double Length = 0;

                Length = Distance.PointToPoint(BeamArray.StartPoint, BeamArray.EndPoint);


                if (ProfileName.Contains("H"))
                {
                    LibraryProfileItem currentProfileItem = GetProfile(ProfileName);


                    if (objectToSelect.Count != 0)
                    {
                        CreateText(fileName, ProfileName, Length, objectToSelect, BeamArray.StartPoint, flag, currentProfileItem);
                    }
                }


                #endregion

            }
        }


        /// <summary>
        /// 生成txt
        /// </summary>
        /// <param name="objectToSelect"></param>
        /// <param name="StartPoint"></param>
        /// <param name="flag"></param>
        /// <param name="j"></param>
        /// <param name="sw"></param>
        /// <param name="beamAllList"></param>
        /// <param name="i"></param>
        /// <param name="CurrentL"></param>
        /// <param name="h"></param>
        /// <param name="L"></param>
        public int GetBoltPointList(ArrayList objectToSelect, Tekla.Structures.Geometry3d.Point StartPoint, int flag, int j, StreamWriter sw, List<List<Beam>> beamAllList, int i, double CurrentL, double h, double L, List<Fitting> fittingList)
        {


            int f;
            double radio = 0; ;

            #region 切割面




            List<Fitting> fitList = new List<Fitting>();

            for (int z = 0; z < fittingList.Count; z++)
            {
                if (fittingList[z].Father.Identifier.ToString() == beamAllList[i][j].Identifier.ToString())
                {
                    fitList.Add(fittingList[z]);
                }

            }




            fitList = fitList.OrderBy(m => Distance.PointToPlane(StartPoint, new GeometricPlane(m.Plane.Origin, m.Plane.AxisX, m.Plane.AxisY))).ToList();


            if (fitList.Count != 0)
            {



                #endregion



                foreach (var bolitItem in objectToSelect)
                {
                    #region

                    List<List<Tekla.Structures.Geometry3d.Point>> pointL = new List<List<Tekla.Structures.Geometry3d.Point>>();

                    List<Tekla.Structures.Geometry3d.Point> pointList = new List<Tekla.Structures.Geometry3d.Point>();
                    List<Tekla.Structures.Geometry3d.Point> pointList1 = new List<Tekla.Structures.Geometry3d.Point>();
                    List<Tekla.Structures.Geometry3d.Point> pointList2 = new List<Tekla.Structures.Geometry3d.Point>();
                    BoltArray bolitArray = bolitItem as BoltArray;

                    radio = bolitArray.BoltSize;

                    ArrayList pointArrayList = bolitArray.BoltPositions;

                    foreach (var pt in pointArrayList)
                    {
                        Tekla.Structures.Geometry3d.Point point = pt as Tekla.Structures.Geometry3d.Point;
                        pointList.Add(point);

                    }

                    pointList = pointList.OrderBy(m => Distance.PointToPlane(m, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY))).ToList();
                    for (int x = 0; x < pointList.Count; x++)
                    {
                        if (x < pointList.Count / 2)
                        {
                            pointList1.Add(pointList[x]);
                        }
                        else
                        {
                            pointList2.Add(pointList[x]);
                        }

                    }

                    pointL.Add(pointList1);
                    pointL.Add(pointList2);




                    #endregion


                    #region 



                    for (int x = 0; x < pointL.Count; x++)
                    {
                        for (int y = 0; y < pointL[x].Count; y++)
                        {
                            flag++;

                            if (j != 0)
                            {
                                double sss = Distance.PointToPlane(pointL[x][y], new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY));
                                sw.WriteLine("[" + flag + "]" + "X:" + (Math.Round(sss + L + 3, 2)) + " Y:********* V:  " + Math.Round(Math.Abs(pointL[x][y].Z - beamAllList[i][j].StartPoint.Z), 2) + " B:********* T1:***** T2:D" + Math.Round(bolitArray.BoltSize, 2) + "  T3:***** PAT:");
                            }
                            else
                            {

                                double sss = Math.Round(Distance.PointToPlane(pointL[x][y], new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2);
                                sw.WriteLine("[" + flag + "]" + "X:" + sss + " Y:********* V:  " + Math.Round(Math.Abs(pointL[x][y].Z - beamAllList[i][j].StartPoint.Z), 2) + " B:********* T1:***** T2:D" + Math.Round(bolitArray.BoltSize, 2) + "  T3:***** PAT:");
                            }




                        }
                    }




                    #endregion

                }
                if (j != beamAllList[i].Count - 1)
                {
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(CurrentL + 1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 - 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(CurrentL + 1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 + 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                }

            }
          



            f = flag;

            return f;

        }


        public void CreatePyText(List<List<Beam>> beamAllList)
        {
            if (Directory.Exists(PyfilePath) == false)
            {

                Directory.CreateDirectory(PyfilePath);

            }

 
            if (!File.Exists(PyfilePath + @"\input_section_table.txt"))
            {
                FileStream fil = new FileStream(PyfilePath + @"\input_section_table.txt", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);

                for (int i = 0; i < beamAllList.Count; i++)
                {
                    List<List<Beam>> beamList = GetNewList(beamAllList[i]);

                    for (int x = 0; x < beamList.Count; x++)
                    {
                        #region

                        double length = 0;
                        beamList[x][0].GetReportProperty("LENGTH", ref length);

                        string NameStr = "";

                        for (int y = 0; y < beamList[x].Count; y++)
                        {

                            if (y != beamList[x].Count - 1)
                            {
                                NameStr += beamList[x][y].Identifier.ToString() + ",";
                            }
                            else
                            {
                                NameStr += beamList[x][y].Identifier.ToString();
                            }

                        }
                        sw.WriteLine(beamList[x][0].Profile.ProfileString + " " + Math.Round(length) + " " + NameStr);

                        #endregion

                    }
                }


                sw.Close();
                fil.Close();


            }
            else
            {
                FileStream fil = new FileStream(PyfilePath + @"\input_section_table.txt", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fil);

                for (int i = 0; i < beamAllList.Count; i++)
                {
                    List<List<Beam>> beamList = GetNewList(beamAllList[i]);

                    for (int x = 0; x < beamList.Count; x++)
                    {

                        #region

                        double length = 0;
                        beamList[x][0].GetReportProperty("LENGTH", ref length);

                        string NameStr = "";

                        for (int y = 0; y < beamList[x].Count; y++)
                        {

                            if (y != beamList[x].Count - 1)
                            {
                                NameStr += beamList[x][y].Identifier.ToString() + ",";
                            }
                            else
                            {
                                NameStr += beamList[x][y].Identifier.ToString();
                            }

                        }
                        sw.WriteLine(beamList[x][0].Profile.ProfileString + " " + Math.Round(length) + " " + NameStr);

                        #endregion


                    }
                }


                sw.Close();
                fil.Close();
            }





        }

        public List<List<Beam>> GetNewList(List<Beam> list)
        {

            List<List<Beam>> beamAllList = new List<List<Beam>>();

            for (int i = 0; i < list.Count; i++)
            {
                double length = 0;
                list[i].GetReportProperty("LENGTH", ref length);

                List<Beam> beamL = new List<Beam>();

                int Flag = 0;
                for (int x = 0; x < beamAllList.Count; x++)
                {
                    for (int y = 0; y < beamAllList[x].Count; y++)
                    {
                        double l = 0;
                        beamAllList[x][y].GetReportProperty("LENGTH", ref l);

                        if (Math.Round(length) == Math.Round(l))
                        {
                            Flag++;
                        }
                    }
                }


                if (Flag == 0)
                {
                    for (int j = 0; j < list.Count; j++)
                    {
                        double length1 = 0;
                        list[j].GetReportProperty("LENGTH", ref length1);
                        if (Math.Round(length) == Math.Round(length1))
                        {
                            beamL.Add(list[j]);
                        }

                    }

                    if (beamL.Count != 0)
                    {
                        beamAllList.Add(beamL);
                    }


                }


            }


            return beamAllList;

        }

        public  void RunPythonScript()
        {
            //Process p = new Process();

            //string path = @"C:\套料优化_ortools\import_section_table.py";
            //p.StartInfo.FileName = @"C:\套料优化_ortools\python.exe";
            //string sArguments = path;

            //p.StartInfo.Arguments = sArguments;
            //p.StartInfo.UseShellExecute = false;

            //p.StartInfo.RedirectStandardOutput = true;

            //p.StartInfo.RedirectStandardInput = true;

            //p.StartInfo.RedirectStandardError = true;

            //p.StartInfo.CreateNoWindow = true;

            //p.Start();

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = @"D:\Program Files\GBS_Software\CreateTeklaCNC" + @"\import_section_table.exe";
            psi.UseShellExecute = false;
            psi.WorkingDirectory = @"D:\Program Files\GBS_Software\CreateTeklaCNC";
            psi.CreateNoWindow = true;

            

           // Process p = Process.Start(@"Y:\数字化课题\数据库\py库调用\套料优化ortools\import_section_table.exe");
            Process p = Process.Start(psi);

            //processForm form = new TestTekla.processForm();
            //form.ShowDialog();


            p.WaitForExit();
          //  form.Close();

        }

        public List<List<string>> ReadTxt(string PyfilePath)
        {
            FileStream file = new FileStream(PyfilePath + @"\out_member_group.txt", FileMode.Open, FileAccess.Read);

            StreamReader sr = new StreamReader(file);
            List<string> GroupList = new List<string>();

            string newLine;
            while ((newLine = sr.ReadLine()) != null)
            {
                if (newLine.Contains("ID_Group"))
                {


                    GroupList.Add(newLine);
                }
            }

            List<List<string>> IdList = new List<List<string>>();



            for (int i = 0; i < GroupList.Count; i++)
            {
                Regex rex = new Regex(@"\(.*?\)");
                string SS = GroupList[i].Split(':')[1];

                MatchCollection ms = rex.Matches(SS);



                foreach (Match m in ms)
                {
                    List<string> idL = new List<string>();

                    string[] StrId = m.Value.Substring(1, m.Value.Length - 2).Split(',');

                    for (int j = 0; j < StrId.Length; j++)
                    {
                        idL.Add(StrId[j]);
                    }

                    IdList.Add(idL);
                }


            }

            return IdList;

        }

        public List<List<string>> ReadDb(string PyfilePath,string ExcelPath)
        {


           List< List<string>> IdList = new List<List<string>>();
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                DataTable table;
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText= "SELECT*FROM Output_table";



                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

                cmd.ExecuteNonQuery();

              
                WriteExcelSecond(table, ExcelPath);


                for (int i = 0; i < table.Rows.Count; i++)
                {
                    List<string> IL = new List<string>();


                    for (int j = 0; j < table.Rows[i].ItemArray[5].ToString().Split(',').Length; j++)
                    {
                        IL.Add(table.Rows[i].ItemArray[5].ToString().Split(',')[j]);
                    }

                    IdList.Add(IL);


                }


            }

         
            return IdList;

        }


        public List<List<Beam>> GetGroupBeam(List<List<string>> GroupIdList, ArrayList objectToSelectBeam)
        {

            List<List<Beam>> beamList = new List<List<Beam>>();


            for (int i = 0; i < GroupIdList.Count; i++)
            {

                List<Beam> bList = new List<Beam>();
                for (int j = 0; j < GroupIdList[i].Count; j++)
                {

                    foreach (var item in objectToSelectBeam)
                    {
                        

                        Beam beam = item as Beam;
                        if (beam != null)
                        {
                            if (beam.Identifier.ToString() == GroupIdList[i][j])
                            {
                                bList.Add(beam);
                            }
                        }

                    }
                }

                beamList.Add(bList);
            }

            return beamList;
        }

       

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fDialog = new FolderBrowserDialog();

            DialogResult result = fDialog.ShowDialog();

            if(result==System.Windows.Forms.DialogResult.Cancel)
            {
                
                return;
            }


            PyfilePath = fDialog.SelectedPath.Trim();

        }


        public  void WriteExcel(List<List<Beam>> beamL,string fileP)
        {

            if (Directory.Exists(fileP) == false)
            {

                Directory.CreateDirectory(fileP);

            }


            IWorkbook wb = new XSSFWorkbook();
 

            ISheet sheet = wb.CreateSheet("套料表");


            #region 表头

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("序号");



         
            row.CreateCell(1).SetCellValue("材质");
            row.CreateCell(2).SetCellValue("文件名");
            row.CreateCell(3).SetCellValue("数量");
            row.CreateCell(4).SetCellValue("切割规格");
            row.CreateCell(5).SetCellValue("长度");

            row.CreateCell(6).SetCellValue("零件号");
            row.CreateCell(7).SetCellValue("单重");

            #endregion

            int firstRow = 1;

            int Flag = 0;
            for (int i = 0; i < beamL.Count; i++)
            {
                string fileName = "";
                #region 获得文件名
                for (int x = 0; x < beamL[i].Count; x++)
                {

                    if (x != beamL[i].Count - 1)
                    {
                        fileName += beamL[i][x].Identifier + "-";
                    }
                    else
                    {
                        fileName += beamL[i][x].Identifier;
                    }



                }

                #endregion


             if(i!=0)
                {
                    sheet.CreateRow(++firstRow).CreateCell(1).SetCellValue("NCX_" + fileName);
                }
             else
                {
                    sheet.CreateRow(firstRow).CreateCell(1).SetCellValue("NCX_" + fileName);
                }
               
                

                for (int j = 0; j < beamL[i].Count; j++)
                {
                    firstRow++;
                    Flag++;
                    string profileName = beamL[i][j].Profile.ProfileString;

                    double length = 0;
                    beamL[i][j].GetReportProperty("LENGTH", ref length);

                    double weight = 0;
                    beamL[i][j].GetReportProperty("WEIGHT", ref weight);

                    

                    string materi = beamL[i][j].Material.MaterialString;


                    #region 
                    IRow rowChild = sheet.CreateRow(firstRow);

                    rowChild.CreateCell(0).SetCellValue(Flag);
                    rowChild.CreateCell(1).SetCellValue(materi);

                    rowChild.CreateCell(2).SetCellValue(profileName + "&&" + "12000");
                    rowChild.CreateCell(3).SetCellValue("1");
                    rowChild.CreateCell(4).SetCellValue(profileName);
                    rowChild.CreateCell(5).SetCellValue(length);
                    rowChild.CreateCell(6).SetCellValue(beamL[i][j].Identifier.ToString());
                    rowChild.CreateCell(7).SetCellValue(weight);
                    #endregion

                }

            }

            using (FileStream fs1 = new FileStream(fileP + "\\" + "套料材料表" + ".xlsx", FileMode.Create))
            {
                wb.Write(fs1);
                fs1.Close();
            }




        }

        public void WriteExcelSecond(DataTable table,string fileP)
        {
            if (Directory.Exists(fileP) == false)
            {

                Directory.CreateDirectory(fileP);

            }



            IWorkbook wb = new XSSFWorkbook();


            ISheet sheet = wb.CreateSheet("套料表");


            #region 表头

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("序号");

            


            row.CreateCell(1).SetCellValue("材质");
            row.CreateCell(2).SetCellValue("文件名");
            sheet.SetColumnWidth(2, 30 * 256);

            row.CreateCell(3).SetCellValue("数量");
            row.CreateCell(4).SetCellValue("切割规格");

            sheet.SetColumnWidth(4, 30 * 256);

            row.CreateCell(5).SetCellValue("长度");

            row.CreateCell(6).SetCellValue("零件号");
            row.CreateCell(7).SetCellValue("单重");
            row.CreateCell(8).SetCellValue("利用率");
            row.CreateCell(9).SetCellValue("余料长度");
            #endregion

            int firstRow = 1;

            int Flag = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {

                int count = table.Rows[i].ItemArray[5].ToString().Split(',').Length;
                string fileName = table.Rows[i].ItemArray[5].ToString().Replace(',', '-');

               

                if (i != 0)
                {

                    IRow r = sheet.CreateRow(++firstRow);
                    r.CreateCell(1).SetCellValue("NCX_" + fileName);
                    r.CreateCell(8).SetCellValue(table.Rows[i].ItemArray[2].ToString());
                    r.CreateCell(9).SetCellValue(table.Rows[i].ItemArray[3].ToString());
                }
                else
                {
                    IRow r = sheet.CreateRow(firstRow);
                    r.CreateCell(1).SetCellValue("NCX_" + fileName);
                    r.CreateCell(8).SetCellValue(table.Rows[i].ItemArray[2].ToString());
                    r.CreateCell(9).SetCellValue(table.Rows[i].ItemArray[3].ToString());
                   
                  
                }


               

                if (count != 1)
                {


                    for (int j = 0; j < count; j++)
                    {
                        Flag++;
                        firstRow++;
                        IRow rowChild = sheet.CreateRow(firstRow);

                        rowChild.CreateCell(0).SetCellValue(Flag);
                        rowChild.CreateCell(1).SetCellValue(table.Rows[i].ItemArray[6].ToString().Split(',')[j]);

                        rowChild.CreateCell(2).SetCellValue(table.Rows[i].ItemArray[0].ToString() + "&&" + "12000");
                        rowChild.CreateCell(3).SetCellValue("1");
                        rowChild.CreateCell(4).SetCellValue(table.Rows[i].ItemArray[0].ToString());
                        rowChild.CreateCell(5).SetCellValue(table.Rows[i].ItemArray[4].ToString().Split(',')[j]);
                        rowChild.CreateCell(6).SetCellValue(table.Rows[i].ItemArray[5].ToString().Split(',')[j]);
                        rowChild.CreateCell(7).SetCellValue(table.Rows[i].ItemArray[7].ToString().Split(',')[j]);
                    }


                }
                else
                {
                    firstRow++;
                    Flag++;
                    IRow rowChild = sheet.CreateRow(firstRow);

                    rowChild.CreateCell(0).SetCellValue(Flag);
                    rowChild.CreateCell(1).SetCellValue(table.Rows[i].ItemArray[6].ToString());

                    rowChild.CreateCell(2).SetCellValue(table.Rows[i].ItemArray[0].ToString() + "&&" + "12000");
                    rowChild.CreateCell(3).SetCellValue("1");
                    rowChild.CreateCell(4).SetCellValue(table.Rows[i].ItemArray[0].ToString());
                    rowChild.CreateCell(5).SetCellValue(table.Rows[i].ItemArray[4].ToString());
                    rowChild.CreateCell(6).SetCellValue(table.Rows[i].ItemArray[5].ToString());
                    rowChild.CreateCell(7).SetCellValue(table.Rows[i].ItemArray[7].ToString());

                }

                #region 
               
                #endregion



            }

            using (FileStream fs1 = new FileStream(fileP + "\\" + "套料材料表" + ".xlsx", FileMode.Create))
            {
                wb.Write(fs1);
                fs1.Close();
            }




        }

        public void AddDataToSqlite(List<List<Beam>> beamAllList)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath+ @"\Tekla_NCX_database.db"))
            {

                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText = " delete from Import_table";
                cmd.ExecuteNonQuery();

                for (int i = 0; i < beamAllList.Count; i++)
                {

                    List<List<Beam>> beamList = GetNewList(beamAllList[i]);


                    for (int x = 0; x < beamList.Count; x++)
                    {
                        #region

                        cmd.CommandText = "Insert into Import_table values(@Section,@Material,@Weight,@Length,@ID_list)";
                        double length = 0;
                        beamList[x][0].GetReportProperty("LENGTH", ref length);

                        double weight = 0;
                        beamList[x][0].GetReportProperty("WEIGHT", ref weight);


                        string NameStr = "";
                        for (int y = 0; y < beamList[x].Count; y++)
                        {

                            if (y != beamList[x].Count - 1)
                            {
                                NameStr += beamList[x][y].Identifier.ToString() + ",";
                            }
                            else
                            {
                                NameStr += beamList[x][y].Identifier.ToString();
                            }

                        }
                        cmd.Parameters.AddWithValue("@Section", beamList[x][0].Profile.ProfileString);
                        cmd.Parameters.AddWithValue("@Material", beamList[x][0].Material.MaterialString);
                        cmd.Parameters.AddWithValue("@Weight", Math.Round(weight,2));
                        cmd.Parameters.AddWithValue("@Length", Math.Round(length,2));
                        cmd.Parameters.AddWithValue("@ID_list", NameStr);

                        cmd.ExecuteNonQuery();
                        #endregion

                    }




                }

            }

        }




       

    }



}
