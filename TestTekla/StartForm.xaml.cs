using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
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
using Tekla.Structures.Catalogs;
using System.Diagnostics;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Tekla.Structures.Model;
using System.Collections;
using System.IO;
using Tekla.Structures.Geometry3d;

namespace TestTekla
{
    /// <summary>
    /// StartForm.xaml 的交互逻辑
    /// </summary>
    public partial class StartForm : Window
    {
        public StartForm()
        {
            InitializeComponent();
        }

        public string filePath;
        public  List<Fitting> fittingList = new List<Fitting>();

        private void button_Click(object sender, RoutedEventArgs e)
        {


            InformationForm form = new TestTekla.InformationForm(textBox.Text);
            form.ShowDialog();
        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            List<Beam> beamList = GetTeklaModule();

           
            if (beamList.Count != 0 && fittingList.Count != 0)
            {



                CreateAllText(beamList, fittingList);
            }
            else
            {
                MessageBox.Show("模型选择有问题");
            }


        }


        public List<Beam> GetTeklaModule()
        {
            Model myModel = new Model();

            filePath = myModel.GetInfo().ModelPath + @"\NCX文件\";

            List<List<Beam>> beamAllList = new List<List<Beam>>();


            List<Beam> objectToSelectBeam = new List<Beam>();


            ModelObjectEnumerator myEnumFitting = myModel.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.FITTING);
         
         
            while (myEnumFitting.MoveNext())
            {
                fittingList.Add(myEnumFitting.Current as Fitting);
            }


            ModelObjectEnumerator selectObjects = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();

            if (selectObjects.GetSize() > 0)
            {
                while (selectObjects.MoveNext())
                {
                    Beam b = selectObjects.Current as Beam;

                   

                    if (b!=null&& b.Profile.ProfileString.Contains("H"))
                    {
                        string partPos = "";

                        b.GetReportProperty("PART_POS", ref partPos);

                        b.PartNumber.Prefix = partPos;

                        objectToSelectBeam.Add(b);
                    }
                    
                }
            }



            objectToSelectBeam.Where((x, i) => objectToSelectBeam.FindIndex(z => z.PartNumber.Prefix == x.PartNumber.Prefix) == i);

            //foreach (var item in objectToSelectBeam)
            //{
                #region 
                //List<Beam> beamList = new List<Beam>();
                //Beam beam = item as Beam;


                //string partPos = "";

                //beam.GetReportProperty("PART_POS", ref partPos);

                //foreach (var itemSecond in objectToSelectBeam)
                //{
                //    Beam beam1 = itemSecond as Beam;


                //    string partP = "";

                //    beam1.GetReportProperty("PART_POS", ref partP);


                //    if (partP == partPos)
                //    {
                //        beamList.Add(beam1);
                //    }



                //}

                //int T = 0;
                //for (int i = 0; i < beamList.Count; i++)
                //{
                // List<Beam> bL=   beamAllList.Find(m => m.Contains(beamList[i]));
                //    if (bL != null)
                //    {


                //        if (bL.Count != 0)
                //        {
                //            T++;
                //        }
                //    }
                   
                //}





                ////int T = 0;
                ////for (int i = 0; i < beamAllList.Count; i++)
                ////{

                ////    for (int j = 0; j < beamAllList[i].Count; j++)
                ////    {
                ////        if (beamList.Contains(beamAllList[i][j]))
                ////        {
                ////            T++;
                ////        }
                ////    }

                ////}
                //if (T == 0)
                //{
                //    if (beamList.Count != 0)
                //    {
                //        beamAllList.Add(beamList);
                //    }

                //}



                #endregion

          //  }



            return objectToSelectBeam;


        }



        public void CreateAllText(List<Beam> beamAllList, List<Fitting> fittingList)
        {

            if (Directory.Exists(filePath) == false)
            {

                Directory.CreateDirectory(filePath);

            }

            try
            {

                MonitorCheck form = new MonitorCheck();
                form.Show();

                int T = 0;
                for (int i = 0; i < beamAllList.Count; i++)
                {
                    string ProString = beamAllList[i].Profile.ProfileString.Replace('*', 'X');
                    double Len = 0;
                    string partPos = "";
                    beamAllList[i].GetReportProperty("LENGTH", ref Len);

                    beamAllList[i].GetReportProperty("PART_POS", ref partPos);

                   // partPos = partPos.Replace('?','1');

                    LibraryProfileItem currentProfileItem = GetProfile(beamAllList[i].Profile.ProfileString);


                    if (currentProfileItem != null)
                    {


                        if (!File.Exists(filePath + "NCX-" + ProString + "-" + Math.Round(Len) + "-" + partPos + ".txt"))
                        {
                           
                           

                            T++;
                            form.SetTextMesssage(T, beamAllList.Count);//进度条函数
                            FileStream fil;
                            StreamWriter sw;
                            #region                    
                            try
                            {

                                 fil = new FileStream(filePath + "NCX-" + ProString + "-" + Math.Round(Len) + "-" + partPos + ".txt", FileMode.Create, FileAccess.Write);

                                 sw = new StreamWriter(fil);
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }

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


                            sw.WriteLine("WORK SIZE :" + "L=" + Math.Round(Len, 2) * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + " TYPE=H" + " FNE:NC1");

                            #endregion
                            int flag = 0;


                            #region

                            ArrayList objectToSelect = new ArrayList();
                            ModelObjectEnumerator myEnum = beamAllList[i].GetBolts();

                            while (myEnum.MoveNext())
                            {
                                objectToSelect.Add(myEnum.Current as BoltArray);
                            }


                            if (objectToSelect.Count != 0)
                            {

                                GetBoltPointList(objectToSelect, beamAllList[i], flag, sw, beamAllList, i, h, fittingList);

                            }
                            else
                            {

                                Tag++;
                                sw.Dispose();
                                fil.Dispose();
                                continue;

                            }


                            #endregion

                            sw.Close();
                            fil.Close();

                            #endregion

                        }
                        else
                        {
                            T++;
                            form.SetTextMesssage(T, beamAllList.Count);//进度条函数

                            FileStream fil;
                            StreamWriter sw;
                            #region
                            try
                            {


                                 fil = new FileStream(filePath + "NCX-" + ProString + "-" + Math.Round(Len) + "-" + partPos + ".txt", FileMode.Create, FileAccess.Write);

                                 sw = new StreamWriter(fil);
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }

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


                            sw.WriteLine("WORK SIZE :" + "L=" + Math.Round(Len, 2) * 1000 + " W1=" + h * 1000 + " T1=" + s * 1000 + " W2=" + b * 1000 + " T2=" + t * 1000 + " W3=" + b * 1000 + " T3=" + t * 1000 + "TYPE=H" + " FNE:NC1");



                            #endregion
                            int flag = 0;

                            #region

                            ArrayList objectToSelect = new ArrayList();
                            ModelObjectEnumerator myEnum = beamAllList[i].GetBolts();

                            while (myEnum.MoveNext())
                            {
                                objectToSelect.Add(myEnum.Current as BoltArray);
                            }


                            if (objectToSelect.Count != 0)
                            {

                                GetBoltPointList(objectToSelect, beamAllList[i], flag, sw, beamAllList, i, h, fittingList);

                            }
                            else
                            {

                                Tag++;
                                sw.Dispose();
                                fil.Dispose();
                                continue;

                            }


                            #endregion


                            sw.Close();
                            fil.Close();
                            #endregion
                        }


                    }
                }

            }catch(Exception ex )
            {
                MessageBox.Show(ex.Message);
            }



        }

        /// <summary>
        /// 获得参数信息
        /// </summary>
        /// <param name="profileName"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 将数据写入txt
        /// </summary>
        /// <param name="objectToSelect"></param>
        /// <param name="beam"></param>
        /// <param name="flag"></param>
        /// <param name="sw"></param>
        /// <param name="beamAllList"></param>
        /// <param name="i"></param>
        /// <param name="h"></param>
        /// <param name="fittingList"></param>
        public void GetBoltPointList(ArrayList objectToSelect, Beam beam, int flag, StreamWriter sw, List<Beam> beamAllList, int i, double h, List<Fitting> fittingList)
        {



            double radio = 0;
            double tolerance = 0;
            #region 切割面


            List<Fitting> fitList = new List<Fitting>();

            for (int z = 0; z < fittingList.Count; z++)
            {
                if (fittingList[z].Father.Identifier.ToString() == beamAllList[i].Identifier.ToString())
                {
                    fitList.Add(fittingList[z]);
                }
            }

            #endregion
            Solid solid = beamAllList[i].GetSolid();         
            Tekla.Structures.Geometry3d.Point p1 = solid.MinimumPoint;

        
            fitList = fitList.OrderBy(m => Distance.PointToPlane(beam.StartPoint, new GeometricPlane(m.Plane.Origin, m.Plane.AxisX, m.Plane.AxisY))).ToList();

            double dis=0;
            if (fitList.Count != 0)
            {
                 dis = Math.Round(Distance.PointToPlane(p1, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2);
            }

                

            if (fitList.Count != 0)
            {


                foreach (var bolitItem in objectToSelect)
                {

                    if (bolitItem != null)
                    {

                        #region

                        List<List<Tekla.Structures.Geometry3d.Point>> pointL = new List<List<Tekla.Structures.Geometry3d.Point>>();

                        List<Tekla.Structures.Geometry3d.Point> pointList = new List<Tekla.Structures.Geometry3d.Point>();
                        List<Tekla.Structures.Geometry3d.Point> pointList1 = new List<Tekla.Structures.Geometry3d.Point>();
                        List<Tekla.Structures.Geometry3d.Point> pointList2 = new List<Tekla.Structures.Geometry3d.Point>();
                        BoltArray bolitArray = bolitItem as BoltArray;

                        radio = bolitArray.BoltSize;


                        tolerance = bolitArray.Tolerance;

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

                                double sss = Math.Round(Distance.PointToPlane(pointL[x][y], new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2)- dis;
                                // double sss = Math.Round(Distance.PointToPlane(pointL[x][y], gPlan), 2);
                                // double sss = Math.Round(Distance.PointToLine(pointL[x][y], line1), 2);
                                sw.WriteLine("[" + flag + "]" + "X:" + sss + " Y:********* V:  " + Math.Round(Math.Abs(pointL[x][y].Z - (beamAllList[i].StartPoint.Z - h)), 2) + " B:********* T1:***** T2:D" + Math.Round(bolitArray.BoltSize + tolerance, 2) + "  T3:***** PAT:");


                            }
                        }

                        #endregion
                    }

                }

            }


        }


    }
}
