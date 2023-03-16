using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlatformCollect
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string DbPath;

        private void btnDb_Click(object sender, EventArgs e)
        {

            OpenFileDialog fileDialog = new OpenFileDialog();

            //选择文件标题
            fileDialog.Title = "请选择Db文件";

            //设置过滤器
            fileDialog.Filter = "所有文件(*.db)|*.ydb";

           

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {

                DbPath = fileDialog.FileName;
            }
        }

        private void btnRevit_Click(object sender, EventArgs e)
        {

            List<RegistryKey> List = new List<RegistryKey>();

            // List.Add(Registry.ClassesRoot);
            // List.Add(Registry.CurrentConfig);
            // List.Add(Registry.CurrentUser);
            List.Add(Registry.LocalMachine);
            // List.Add(Registry.PerformanceData);
            // List.Add(Registry.Users);


            // Dictionary<string, string> listpath = new Dictionary<string, string>();
            string softWarePath = null;

            Dictionary<string, string> softwar = new Dictionary<string, string>();


            string SubKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

            foreach (RegistryKey reg in List)
            {

                using (RegistryKey registryKey1 = reg.OpenSubKey(SubKeyName))
                {
                    if (registryKey1 == null)

                        continue;
                    if (registryKey1.GetSubKeyNames() == null)
                        continue;

                    // RegistryValueKind regVaulekind = registryKey1.GetValueKind(string.Empty);
                    string[] KeyNames = registryKey1.GetSubKeyNames();
                    foreach (string keyName in KeyNames)
                    {
                        using (RegistryKey registryKey2 = registryKey1.OpenSubKey(keyName))
                        {
                            if (registryKey2 == null)
                                continue;
                            string softwareName = registryKey2.GetValue("DisplayName", "").ToString();
                            string InstallLocation = registryKey2.GetValue("InstallLocation", "").ToString();

                            if (!string.IsNullOrEmpty(InstallLocation) && !string.IsNullOrEmpty(softwareName))
                            {
                                if (!softwar.ContainsKey(softwareName))
                                    softwar.Add(softwareName, InstallLocation);

                            }
                        }


                    }
                }

            }

            if (softwar.Count > 0)
            {
                foreach (string softwareName in softwar.Keys)
                {
                    if (softwareName.Equals("Autodesk Revit 2018"))
                    {
                        softWarePath = softwar[softwareName];
                    }

                }
            }




            System.Diagnostics.Process.Start(softWarePath + "Revit.exe");
            Close();
        }
    }
}
