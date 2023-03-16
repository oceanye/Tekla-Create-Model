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

namespace TestTekla
{
    /// <summary>
    /// DetailInformatibForm.xaml 的交互逻辑
    /// </summary>
    public partial class DetailInformatibForm : Window
    {
        public DetailInformatibForm(List<SqliteData> sqlD)
        {

            sqlList = sqlD;



            InitializeComponent();
        }

        public List<SqliteData> sqlList = new List<SqliteData>();

       
        public string PyfilePath = @"D:\Program Files\GBS_Software\CreateTeklaCNC";

        public List<CheckBox> headerChecks = new List<CheckBox>();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            Title = sqlList[0].ProfileName + " 详细信息";
            dataGrid.AutoGenerateColumns = false;
            dataGrid.ItemsSource = sqlList;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            double MateriAllNumber = ((sender as Button).BindingGroup.Items[0] as SqliteData).MateriAllNumber;
            string profileName = ((sender as Button).BindingGroup.Items[0] as SqliteData).ProfileName;

            string IdList= ((sender as Button).BindingGroup.Items[0] as SqliteData).IdList;

            DeleteDbById(IdList);

            DeleteDbByProfile(PyfilePath, MateriAllNumber);
            List<SqliteData> sqlList= ReadDbByProfile(PyfilePath, profileName);

            dataGrid.AutoGenerateColumns = false; 
            dataGrid.ItemsSource = sqlList;

        }


        public void DeleteDbByProfile(string PyfilePath, double  number)
        {
     
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
            {

                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;

                cmd.CommandText = "delete from Output_table Where Bar_group=@Bar_group";

                cmd.Parameters.AddWithValue("@Bar_group", number);

                cmd.ExecuteNonQuery();

            }

        }


        public void DeleteDbById(string IdList)
        {
            for (int i = 0; i < IdList.Split(',').Length; i++)
            {
                using (SQLiteConnection conn = new SQLiteConnection("Data Source=" + PyfilePath + @"\Tekla_NCX_database.db"))
                {

                    conn.Open();

                    SQLiteCommand cmd = new SQLiteCommand();

                    cmd.Connection = conn;





                    string sqlImport = "SELECT*FROM Import_table where ID_list Like '%" + IdList.Split(',')[i] + "%'";
                    DataTable dataTable = GetSqlite(sqlImport);

                    List<string> IList = new List<string>();
                    for (int j = 0; j < dataTable.Rows.Count; j++)
                    {

                        List<string> IdL = dataTable.Rows[j].ItemArray[4].ToString().Split(',').ToList();

                        for (int x = 0; x < IdL.Count; x++)
                        {
                            if (IdL[x] != IdList.Split(',')[i])
                            {
                                IList.Add(IdL[x]);
                            }
                        }



                    }

                    string IdListNew = "";

                    for (int y = 0; y < IList.Count; y++)
                    {
                        if (y == IList.Count - 1)
                        {
                            IdListNew += IList[y];
                        }
                        else
                        {
                            IdListNew += IList[y] + ",";
                        }

                    }

                    cmd.CommandText = "UPDATE Import_table SET ID_list=@ID_list where ID_list Like '%" + IdList.Split(',')[i] + "%'";

                    cmd.Parameters.AddWithValue("@ID_list", IdListNew);
                   

                    int xxx = cmd.ExecuteNonQuery();
                }



            }

        }


        public List<SqliteData> ReadDbByProfile(string PyfilePath, string profile)
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
          
            List<SqliteData> sqlList = new List<SqliteData>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                SqliteData sqlD = new SqliteData();
                sqlD.ProfileName = table.Rows[i].ItemArray[0].ToString();
                sqlD.MateriAllNumber = Convert.ToDouble(table.Rows[i].ItemArray[1]);
                sqlD.Rate = Convert.ToDouble(table.Rows[i].ItemArray[2]);

                sqlD.WasteLength = Convert.ToDouble(table.Rows[i].ItemArray[3]);
                sqlD.LengthList = table.Rows[i].ItemArray[4].ToString();
                sqlD.IdList = table.Rows[i].ItemArray[5].ToString();               
                sqlD.MateriList = table.Rows[i].ItemArray[6].ToString();
                sqlD.WeightList = table.Rows[i].ItemArray[7].ToString();

                sqlList.Add(sqlD);
            }

            return sqlList;

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

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            CheckBox cb = (CheckBox)sender;

            headerChecks.ForEach(a => a.IsChecked = cb.IsChecked.Value);

        }


        private void CheckBox_Loaded(object sender, RoutedEventArgs e)
        {
            CheckBox cb = (CheckBox)sender;

            headerChecks.Add(cb);

        }

        private void button_Click_1(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < headerChecks.Count; i++)
            {
                if ((headerChecks[i] as CheckBox).IsChecked == true)
                {


                    double MateriAllNumber = ((headerChecks[i] as CheckBox).BindingGroup.Items[0] as SqliteData).MateriAllNumber;
                    string profileName = ((headerChecks[i] as CheckBox).BindingGroup.Items[0] as SqliteData).ProfileName;


                    for (int j = 0; j < sqlList.Count; j++)
                    {
                        if(sqlList[j].MateriAllNumber== MateriAllNumber)
                        {
                            sqlList.Remove(sqlList[j]);
                        }
                    }

                    string IdList = ((headerChecks[i] as CheckBox).BindingGroup.Items[0] as SqliteData).IdList;

                    DeleteDbById(IdList);

                    DeleteDbByProfile(PyfilePath, MateriAllNumber);
                    List<SqliteData> sqlL = ReadDbByProfile(PyfilePath, profileName);

                    dataGrid.AutoGenerateColumns = false;
                    dataGrid.ItemsSource = sqlL;



                }
            }
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {

            


             List<SqliteData> sqlFilterList = new List<SqliteData>();

            if (textBox2.Text != "")
            {


                for (int i = 0; i < sqlList.Count; i++)
                {
                    if (sqlList[i].Rate < Convert.ToDouble(textBox2.Text))
                    {
                        sqlFilterList.Add(sqlList[i]);
                    }
                }
                dataGrid.ItemsSource = sqlFilterList;

                headerChecks.Clear();

             

            }
            else
            {
                MessageBox.Show("请输入利用率");
            }
        }

       
    }
}
