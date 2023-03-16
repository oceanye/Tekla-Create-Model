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
using System.Drawing;
using QRCoder;
using System.Windows.Forms;

namespace TestTekla
{
    /// <summary>
    /// InformationForm.xaml 的交互逻辑
    /// </summary>
    public partial class InformationForm : Window
    {
        public InformationForm(string text)
        {

            //dataGrid.MouseDown += DataGrid_MouseDown;

            toleranceText = text;
            InitializeComponent();
        }


        public string toleranceText;

        private void DataGrid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            throw new NotImplementedException();
        }

        public string PyfilePath = @"D:\Program Files\GBS_Software\CreateTeklaCNC";

        public string filePath;

        public List<System.Windows.Controls.CheckBox> headerChecks = new List<System.Windows.Controls.CheckBox>();

        public string QRPath = @"C:\NCX_Data\";

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            Title = "数据生成中....";
            List<List<Beam>> beamList = GetTeklaModule();

            if(beamList.Count>0&& beamList[0].Count>=1)
            {


                AddDataToSqlit(beamList);//将数据存入数据库

                RunPythonScript();//调用py脚本

                Title = "截面信息统计"; 
                dataGrid.AutoGenerateColumns = false;
                dataGrid.ItemsSource = GetSqlDate();

                string sql= "SELECT*FROM Output_table";
                GetSqlite(sql);

                string sqlImport = "SELECT*FROM Import_table";
                DataTable sqlDate= GetSqlite(sqlImport);

                List<string> IL = new List<string>();

                for (int i = 0; i < sqlDate.Rows.Count; i++)
                {

                    if (sqlDate.Rows[i].ItemArray[4].ToString() != "")
                    {
                        string[] idList = sqlDate.Rows[i].ItemArray[4].ToString().Split(',');

                        for (int j = 0; j < idList.Length; j++)
                        {
                            IL.Add(idList[j]);
                        }
                    }

                }

                textBox.Text = IL.Count.ToString();

                textBox1.Text = GetSqlite(sql).Rows.Count.ToString();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请查看构件是否单一长度或已分组");
                Close();
            }

        }

        public List<List<DataRow>> ReadDb(string PyfilePath)
        {


            List<List<DataRow>> beamAllList = new List<List<DataRow>>();

            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                DataTable table;
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

               

                cmd.CommandText = "SELECT*FROM Output_table";



                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

                cmd.ExecuteNonQuery();


              


                for (int i = 0; i < table.Rows.Count; i++)
                {
                    #region

                    List<DataRow> beamList = new List<DataRow>();
                    string profileName = table.Rows[i].ItemArray[0].ToString();

                    for (int j = 0; j < table.Rows.Count; j++)
                    {

                        string pName= table.Rows[j].ItemArray[0].ToString();

                        if(profileName== pName)
                        {
                            beamList.Add(table.Rows[j]);
                        }
                    }


                    int T = 0;
                    for (int x = 0; x < beamAllList.Count; x++)
                    {

                        for (int y = 0; y < beamAllList[x].Count; y++)
                        {
                            if (beamList.Contains(beamAllList[x][y]))
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

                    #endregion
                }



            }

            return beamAllList;

        }
 

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string profileName  = ((sender as System.Windows.Controls.Button).BindingGroup.Items[0] as SqliteData).ProfileName;

            DataTable table= ReadDbByProfile(PyfilePath, profileName);

            List<SqliteData> sqlList = new List<SqliteData>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                SqliteData sqlD = new SqliteData();
                sqlD.ProfileName = table.Rows[i].ItemArray[0].ToString();
                sqlD.MateriAllNumber = Convert.ToDouble(table.Rows[i].ItemArray[1]);
                sqlD.Rate= Convert.ToDouble(table.Rows[i].ItemArray[2]);
              
                sqlD.WasteLength= Convert.ToDouble(table.Rows[i].ItemArray[3]);
                sqlD.LengthList= table.Rows[i].ItemArray[4].ToString();
                sqlD.IdList = table.Rows[i].ItemArray[5].ToString();
                sqlD.MateriList = table.Rows[i].ItemArray[6].ToString();
                sqlD.WeightList = table.Rows[i].ItemArray[7].ToString();
               


                sqlList.Add(sqlD);
            }

            sqlList = sqlList.OrderBy(m => m.Rate).ToList();

  

            DetailInformatibForm form = new DetailInformatibForm(sqlList);

            form.ShowDialog();
          
        }

        public DataTable ReadDbByProfile(string PyfilePath,string profile)
        {



            DataTable table;
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText = "SELECT*FROM Output_table Where Section=@Section";

                cmd.Parameters.AddWithValue("@Section", profile);

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

                cmd.ExecuteNonQuery();
       



            }

            return table;

        }

        private void button_Click_1(object sender, RoutedEventArgs e)
        {
            dataGrid.AutoGenerateColumns = false;
            dataGrid.ItemsSource = GetSqlDate();


            string sql = "SELECT*FROM Output_table";
            GetSqlite(sql);

            string sqlImport = "SELECT*FROM Import_table";
            DataTable sqlDate = GetSqlite(sqlImport);

            List<string> IL = new List<string>();

            for (int i = 0; i < sqlDate.Rows.Count; i++)
            {

                if (sqlDate.Rows[i].ItemArray[4].ToString() != "")
                {
                    string[] idList = sqlDate.Rows[i].ItemArray[4].ToString().Split(',');

                    for (int j = 0; j < idList.Length; j++)
                    {
                        IL.Add(idList[j]);
                    }
                }

            }

            textBox.Text = IL.Count.ToString();

            textBox1.Text = GetSqlite(sql).Rows.Count.ToString();

        

        }

        public List<SqliteData> GetSqlDate()
        {
            List<List<DataRow>> bList = ReadDb(PyfilePath);

            List<SqliteData> sqlDataList = new List<SqliteData>();

            for (int i = 0; i < bList.Count; i++)
            {
                SqliteData sqlD = new SqliteData();
                double AvRate = 0;

                for (int j = 0; j < bList[i].Count; j++)
                {
                    AvRate += Convert.ToDouble(bList[i][j].ItemArray[2]);
                }
                sqlD.AverageRate = Math.Round(AvRate / bList[i].Count, 2);
                sqlD.ProfileName = bList[i][0].ItemArray[0].ToString();
                sqlD.MaterialName = bList[i][0].ItemArray[6].ToString().Split(',')[0];

                sqlD.Number = bList[i].Count;
                sqlD.Other = "";

                sqlDataList.Add(sqlD);

            }

            return sqlDataList;

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Model myModel = new Model();
            string path = myModel.GetInfo().ModelPath + @"\EXCEL文件\";
          
            AddDataToExcel(PyfilePath, path);//将数据写入EXCEL




            System.Windows.Forms.MessageBox.Show("完成");

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {

            Model myModel = new Model();
            List<Fitting> fittingList = new List<Fitting>();
            ArrayList objectToSelectBeam = new ArrayList();
            ModelObjectEnumerator selectObjects = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();
            ModelObjectEnumerator myEnumFitting = myModel.GetModelObjectSelector().GetAllObjectsWithType(Tekla.Structures.Model.ModelObject.ModelObjectEnum.FITTING);

            filePath = myModel.GetInfo().ModelPath + @"\合并文件\";

            if (selectObjects.GetSize() > 0)
            {
                while (selectObjects.MoveNext())
                {
                    objectToSelectBeam.Add(selectObjects.Current as Beam);
                }
            }

            while (myEnumFitting.MoveNext())
            {
                fittingList.Add(myEnumFitting.Current as Fitting);
            }


            int Flag = 0;
            foreach (var item in objectToSelectBeam)
            {
                Beam beam = item as Beam;
                string partPos = "";

                if (beam != null)
                {



                    beam.GetReportProperty("PART_POS", ref partPos);

                    if (partPos.Contains("?"))
                    {
                        Flag++;

                       



                    }
                }
            }


            if(Flag>0)
            {
                DialogResult result = System.Windows.Forms.MessageBox.Show("组件编号含有非法字符(?),请重新运行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question);

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    Close();
                    //List<List<string>> GroupIdList = ReadSqlite(PyfilePath);


                    //List<List<Beam>> beamAllListSecond = GetGroupBeam(GroupIdList, objectToSelectBeam, myModel);




                    //CreateAllText(beamAllListSecond, fittingList);//分组文件
                }
                //else if (result == System.Windows.Forms.DialogResult.Cancel)
                //{
                //    Close();
                    
                //}
            }
            else
            {
                List<List<string>> GroupIdList = ReadSqlite(PyfilePath);


                List<List<Beam>> beamAllListSecond = GetGroupBeam(GroupIdList, objectToSelectBeam, myModel);




                CreateAllText(beamAllListSecond, fittingList);//分组文件
            }

           
        }


        public void AddDataToSqlit(List<List<Beam>> beamAllList)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
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

                        cmd.CommandText = "Insert into Import_table values(@Section,@Material,@Weight,@Length,@ID_list,@Number_list)";
                        double length = 0;
                        beamList[x][0].GetReportProperty("LENGTH", ref length);

                        double weight = 0;
                        beamList[x][0].GetReportProperty("WEIGHT", ref weight);


                        string partList = "";
                        string NameStr = "";
                        for (int y = 0; y < beamList[x].Count; y++)
                        {

                            string partPos = "";

                            beamList[x][y].GetReportProperty("PART_POS", ref partPos); 

                            if (y != beamList[x].Count - 1)
                            {

                                partList += partPos+",";
                                NameStr += beamList[x][y].Identifier.ToString() + ",";
                            }
                            else
                            {
                                partList += partPos;
                                NameStr += beamList[x][y].Identifier.ToString();
                            }

                        }
                        cmd.Parameters.AddWithValue("@Section", beamList[x][0].Profile.ProfileString);
                        cmd.Parameters.AddWithValue("@Material", beamList[x][0].Material.MaterialString);
                        cmd.Parameters.AddWithValue("@Weight", Math.Round(weight, 2));
                        cmd.Parameters.AddWithValue("@Length", Math.Round(length, 2));
                        cmd.Parameters.AddWithValue("@ID_list", NameStr);
                        cmd.Parameters.AddWithValue("@Number_list", partList);
                        cmd.ExecuteNonQuery();
                        #endregion

                    }




                }

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

        /// <summary>
        /// 获得TEkla模型
        /// </summary>
        /// <returns></returns>
        public List<List<Beam>> GetTeklaModule()
        {
            Model myModel = new Model();

            List<List<Beam>> beamAllList = new List<List<Beam>>();

            ArrayList objectToSelectBeam = new ArrayList();

            List<Fitting> fittingList = new List<Fitting>();


            ModelObjectEnumerator selectObjects = new Tekla.Structures.Model.UI.ModelObjectSelector().GetSelectedObjects();

            if (selectObjects.GetSize() > 0)
            {
                while (selectObjects.MoveNext())
                {
                    objectToSelectBeam.Add(selectObjects.Current as Beam);
                }
            }

            foreach (var item in objectToSelectBeam)
            {
                #region 


                List<Beam> beamList = new List<Beam>();
                Beam beam = item as Beam;

               
                if (beam != null)
                {
                    string commentText = null;
                    beam.GetUserProperty("comment", ref commentText);




                    if (commentText != "机加工文件已生成")
                    {

                        double length = 0;

                        beam.GetReportProperty("LENGTH", ref length);

                   


                        if (length < 12000)
                        {

                            if (beam.Profile.ProfileString.Contains("H"))
                            {

                                foreach (var itemSecond in objectToSelectBeam)
                                {
                                    Beam beam1 = itemSecond as Beam;
                                   

                                    if (beam1 != null )
                                    {
                                        string commentText1 = null;
                                        beam1.GetUserProperty("comment", ref commentText1);
                                        if (commentText1 != "机加工文件已生成")
                                        {

                                            double l = 0;

                                            beam1.GetReportProperty("LENGTH", ref l);
                                            if (l < 12000)
                                            {

                                                if (beam.Profile.ProfileString == beam1.Profile.ProfileString)
                                                {
                                                    beamList.Add(beam1);
                                                }
                                            }
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

                }

                #endregion

            }



            return beamAllList;


        }

        /// <summary>
        /// 将数据库写入EXCEL
        /// </summary>
        /// <param name="PyfilePath"></param>
        /// <param name="ExcelPath"></param>
        public  void AddDataToExcel(string PyfilePath, string ExcelPath)
        {


            List<List<string>> IdList = new List<List<string>>();
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                DataTable table;
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText = "SELECT*FROM Output_table";



                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

                cmd.ExecuteNonQuery();


                WriteExcelSecond(table, ExcelPath);


            }
         

        }


        public void WriteExcelSecond(DataTable table, string fileP)
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
            row.CreateCell(10).SetCellValue("编号");
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
                    r.CreateCell(1).SetCellValue("NCX_"+ table.Rows[i].ItemArray[0].ToString().Replace('*', 'X') + "_" + fileName);

                    r.Height = 80 * 20;

                    #region 二维码

                    int F = firstRow;

                    

                    Bitmap bitmap = CreateQRCode(QRPath + "NCX_" +table.Rows[i].ItemArray[0].ToString().Replace('*', 'X') + "_" + fileName + ".txt", 11, 7, 5, 5, true);

                    byte[] bytes = BitmapToByte(bitmap);

                    int pictureIdx = wb.AddPicture(bytes, PictureType.JPEG);

                    XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();

                    XSSFClientAnchor abchor = new XSSFClientAnchor(70, 11, 0, 0, 11, F, 13, F + 1);
                    XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(abchor, pictureIdx);

                    #endregion



                    r.CreateCell(8).SetCellValue(table.Rows[i].ItemArray[2].ToString());
                    r.CreateCell(9).SetCellValue(table.Rows[i].ItemArray[3].ToString());
                }
                else
                {
                    IRow r = sheet.CreateRow(firstRow);
                    r.CreateCell(1).SetCellValue("NCX_" + table.Rows[i].ItemArray[0].ToString().Replace('*', 'X') + "_" + fileName);


                    r.Height = 80 * 20;

                    #region 二维码

                    int F = firstRow;

                   

                    Bitmap bitmap = CreateQRCode(QRPath + "NCX_" + table.Rows[i].ItemArray[0].ToString().Replace('*', 'X') + "_" + fileName + ".txt", 11, 7, 5, 5, true);

                    byte[] bytes = BitmapToByte(bitmap);

                    int pictureIdx = wb.AddPicture(bytes, PictureType.JPEG);

                    XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();

                    XSSFClientAnchor abchor = new XSSFClientAnchor(70, 11, 0, 0, 11, F, 13, F + 1);
                    XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(abchor, pictureIdx);

                    #endregion

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
                        rowChild.CreateCell(6).SetCellValue(table.Rows[i].ItemArray[8].ToString().Split(',')[j]); // 零件号
                        rowChild.CreateCell(7).SetCellValue(table.Rows[i].ItemArray[7].ToString().Split(',')[j]);
                        if(table.Rows[i].ItemArray[8]!=null)
                        {
                            rowChild.CreateCell(10).SetCellValue(table.Rows[i].ItemArray[5].ToString().Split(',')[j]);// ID号
                        }
                        else
                        {
                            rowChild.CreateCell(10).SetCellValue("");
                        }
                      
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
                    rowChild.CreateCell(6).SetCellValue(table.Rows[i].ItemArray[8].ToString());// 零件号
                    rowChild.CreateCell(7).SetCellValue(table.Rows[i].ItemArray[7].ToString());
                    if (table.Rows[i].ItemArray[8] != null)
                    {
                        rowChild.CreateCell(10).SetCellValue(table.Rows[i].ItemArray[5].ToString());// ID号
                    }
                    else
                    {
                        rowChild.CreateCell(10).SetCellValue("");
                    }
                }

               
            }

            using (FileStream fs1 = new FileStream(fileP + "\\" + "套料材料表" + ".xlsx", FileMode.Create))
            {
                wb.Write(fs1);
                fs1.Close();
            }




        }

        /// <summary>
        /// 读取数据库文件
        /// </summary>
        /// <param name="PyfilePath"></param>
        /// <returns></returns>
        public List<List<string>> ReadSqlite(string PyfilePath)
        {


            List<List<string>> IdList = new List<List<string>>();
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                DataTable table;
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText = "SELECT*FROM Output_table";



                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

                cmd.ExecuteNonQuery();



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

        public List<List<Beam>> GetGroupBeam(List<List<string>> GroupIdList, ArrayList objectToSelectBeam, Model myModel)
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

                                beam.SetUserProperty("comment", "已分组");

                               
                                bList.Add(beam);
                            }
                        }

                    }
                }

                beamList.Add(bList);
            }
            myModel.CommitChanges();
            return beamList;
        }

        /// <summary>
        /// 生成NCX文件
        /// </summary>
        /// <param name="beamAllList"></param>
        /// <param name="fittingList"></param>
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

                string Name = beamAllList[i][0].Profile.ProfileString.Replace('*', 'X');

            for (int j = 0; j < beamAllList[i].Count; j++)
                {

                    double length = 0;
                    beamAllList[i][j].GetReportProperty("LENGTH", ref length);

                    if (j != beamAllList[i].Count - 1)
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


                    if (!File.Exists(filePath + "NCX_" + Name +"_"+ SumString + ".txt"))
                    {
                        T++;
                        form.SetTextMesssage(T, beamAllList.Count);//进度条函数

                        #region 


                        FileStream fil = new FileStream(filePath + "NCX_" + Name + "_" + SumString + ".txt", FileMode.Create, FileAccess.Write);
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

                                for (int m = 0; m < j; m++)
                                {
                                    double le = 0;
                                    beamAllList[i][j - (m + 1)].GetReportProperty("LENGTH", ref le);

                                    L += le;


                                }

                            }


                            for (int n = 0; n < j + 1; n++)
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

                                flag = GetBoltPointList(objectToSelect, beamAllList[i][j],flag, j, sw, beamAllList, i, CurrentL, h, SumLength, fittingList);

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
                        FileStream fil = new FileStream(filePath + "NCX_" + Name + "_" + SumString + ".txt", FileMode.Create, FileAccess.Write);
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

                                flag = GetBoltPointList(objectToSelect, beamAllList[i][j], flag, j, sw, beamAllList, i, CurrentL, h, SumLength, fittingList);


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

        /// <summary>
        /// 获得螺栓组
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
        /// <param name="fittingList"></param>
        /// <returns></returns>
        public int GetBoltPointList(ArrayList objectToSelect, Beam beam, int flag, int j, StreamWriter sw, List<List<Beam>> beamAllList, int i, double CurrentL, double h, double L, List<Fitting> fittingList)
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

            #endregion


            fitList = fitList.OrderBy(m => Distance.PointToPlane(beam.StartPoint, new GeometricPlane(m.Plane.Origin, m.Plane.AxisX, m.Plane.AxisY))).ToList();

            Solid solid = beamAllList[i][j].GetSolid();

            double dis = 0;//梁的端点到切割面的距离
            if (fitList.Count != 0)
            {



                Tekla.Structures.Geometry3d.Point p1 = null;

                double d = Math.Round(Distance.PointToPlane(beam.StartPoint, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2);
                if (Math.Round(Distance.PointToPlane(beam.StartPoint, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2) > 200)
                {
                    p1 = solid.MaximumPoint;//最大位置
                }
                else
                {
                    p1 = solid.MinimumPoint;//最小位置
                }



                 dis = Math.Round(Distance.PointToPlane(p1, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2);

                if(dis>100)
                {
                    if (Math.Round(Distance.PointToPlane(beam.StartPoint, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2) > 200)
                    {
                        p1 = solid.MinimumPoint;
                    }
                    else
                    {
                        p1 = solid.MaximumPoint;
                    }

                    dis = Math.Round(Distance.PointToPlane(p1, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2);
                }

            }
            #region 测试方法
           // Tekla.Structures.Geometry3d.Line line = new Tekla.Structures.Geometry3d.Line(beam.StartPoint, beam.EndPoint);

           // Tekla.Structures.Geometry3d.Point IntersectionPoint =  Intersection.LineToPlane(line, new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY));


          
           // GeometricPlane gPlan = new GeometricPlane(IntersectionPoint, new Tekla.Structures.Geometry3d.Vector(IntersectionPoint.X, IntersectionPoint.Y, IntersectionPoint.Z), new Tekla.Structures.Geometry3d.Vector(IntersectionPoint.X, IntersectionPoint.Y, IntersectionPoint.Z-1000));

          //  GeometricPlane gPlan = new GeometricPlane(IntersectionPoint, new Tekla.Structures.Geometry3d.Vector(10000,0,0));

           
            #endregion

            if (fitList.Count != 0)
            {


                if (j == 0)
                {
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(Convert.ToDouble(toleranceText)+1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 - 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");//起始点位
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(Convert.ToDouble(toleranceText) + 1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 + 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                }

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
                                double sss = Distance.PointToPlane(pointL[x][y], new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY))-dis-Convert.ToDouble( toleranceText);
                                //double sss = Distance.PointToPlane(pointL[x][y], gPlan)- dis;
                               // double sss = Math.Round(Distance.PointToLine(pointL[x][y], line1), 2);
                                sw.WriteLine("[" + flag + "]" + "X:" + (Math.Round(sss + L + 3, 2)) + " Y:********* V:  " + Math.Round(Math.Abs(pointL[x][y].Z - (beamAllList[i][j].StartPoint.Z-h)), 2) + " B:********* T1:***** T2:D" + Math.Round(bolitArray.BoltSize, 2) + "  T3:***** PAT:");
                            }
                            else
                            {
                                double sss = Math.Round(Distance.PointToPlane(pointL[x][y], new GeometricPlane(fitList[0].Plane.Origin, fitList[0].Plane.AxisX, fitList[0].Plane.AxisY)), 2)-dis - Convert.ToDouble(toleranceText);
                               // double sss = Math.Round(Distance.PointToPlane(pointL[x][y], gPlan), 2)- dis;
                               // double sss = Math.Round(Distance.PointToLine(pointL[x][y], line1), 2);
                                sw.WriteLine("[" + flag + "]" + "X:" + sss + " Y:********* V:  " + Math.Round(Math.Abs(pointL[x][y].Z - (beamAllList[i][j].StartPoint.Z - h)), 2) + " B:********* T1:***** T2:D" + Math.Round(bolitArray.BoltSize, 2) + "  T3:***** PAT:");


                            }

                        }
                    }

                    #endregion

                }
                
                if (j != beamAllList[i].Count - 1)
                {
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(CurrentL + 1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 - 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");//中间MD点位
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(CurrentL + 1.5, 2) + " Y:********* V:  " + Math.Round(h / 2 + 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                }else if(j== beamAllList[i].Count - 1)
                {
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(L + 3+ Convert.ToDouble(toleranceText), 2) + " Y:********* V:  " + Math.Round(h / 2 - 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                    sw.WriteLine("[" + (++flag) + "]" + "X:" + Math.Round(L + 3+ Convert.ToDouble(toleranceText), 2) + " Y:********* V:  " + Math.Round(h / 2 + 100, 2) + " B:********* T1:***** T2:MD" + Math.Round((objectToSelect[0] as BoltArray).BoltSize, 2) + "  T3:***** PAT:");
                }

            }

            f = flag;

            return f;

        }

        /// <summary>
        /// 获得截面
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
            if (correntProfileItem == null):
                {

                 }

            return correntProfileItem;

        }

        /// <summary>
        /// 调用py
        /// </summary>
        public void RunPythonScript()
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


        public DataTable GetSqlite(string sql)
        {

            DataTable table = new DataTable();

            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = conn;

                cmd.CommandText = sql;

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                table = ds.Tables[0];

            }

            return table;

        }

        public Bitmap CreateQRCode(string textAddress,int version,int pixel,int icon_size,int icon_border,bool white_edge)
        {
            QRCodeGenerator code_generator = new QRCodeGenerator();
            QRCodeData code_data = code_generator.CreateQrCode(textAddress, QRCodeGenerator.ECCLevel.M, true, true, QRCodeGenerator.EciMode.Utf8, version);

            QRCode code = new QRCode(code_data);


           // Bitmap icon = new Bitmap(icon_path);

            Bitmap bmp = code.GetGraphic(pixel, System.Drawing.Color.Black, System.Drawing.Color.White,null, icon_size, icon_border, white_edge);


           

            return bmp;

        }

       


        private byte[] BitmapToByte(Bitmap bitmap)
        {
            // 1.先将BitMap转成内存流
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            ms.Seek(0, System.IO.SeekOrigin.Begin);
            // 2.再将内存流转成byte[]并返回
            byte[] bytes = new byte[ms.Length];
            ms.Read(bytes, 0, bytes.Length);
            ms.Dispose();
            return bytes;
        }





    }
}
