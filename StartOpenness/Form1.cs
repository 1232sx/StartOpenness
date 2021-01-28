using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;

namespace StartOpenness
{
    public partial class Form1 : Form
    {

        private static TiaPortalProcess _tiaProcess;

        public TiaPortal MyTiaPortal
        {
            get; set;
        }
        public Project MyProject
        {
            get; set;
        }


        public Form1()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }
        /// <summary>
        /// Function which is called in start after initialization of Form1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private static Assembly MyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            if (index == -1)
            {
                return null;
            }
            string name = args.Name.Substring(0, index);

            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey(
                "SOFTWARE\\Siemens\\Automation\\Openness\\15.1\\PublicAPI\\15.1.0.0");

            if(filePathReg == null)
                return null;

            object oRegKeyValue = filePathReg.GetValue(name);
            if (oRegKeyValue != null)
            {
                string filePath = oRegKeyValue.ToString();

                string path = filePath;
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return Assembly.LoadFrom(fullPath);
                }
            }
           
            return null;
        }
        /// <summary>
        /// SIEMENS function - event for a START TIA button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StartTIA(object sender, EventArgs e)
        {
            if (rdb_WithoutUI.Checked == true)
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                txt_Status.Text = "TIA Portal started without user interface";
                _tiaProcess = TiaPortal.GetProcesses()[0];
            }
            else
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                txt_Status.Text = "TIA Portal started with user interface";
            }

            btn_SearchProject.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_Start.Enabled = false;

        }
        /// <summary>
        /// SIEMENS function - event for a DISPOSE TIA button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DisposeTIA(object sender, EventArgs e)
        {
            MyTiaPortal.Dispose();
            txt_Status.Text = "TIA Portal disposed";
            btn_Start.Enabled = true;
            btn_Dispose.Enabled = false;
            btn_CloseProject.Enabled = false;
            btn_SearchProject.Enabled = false;
            btn_CompileHW.Enabled = false;
            btn_Save.Enabled = false;

        }
        /// <summary>
        /// SIEMENS function - event for a OPEN PROJECT button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SearchProject(object sender, EventArgs e)
        {
            OpenFileDialog fileSearch = new OpenFileDialog();
            fileSearch.Filter = "*.ap15_1|*.ap15_1";
            fileSearch.RestoreDirectory = true;
            fileSearch.ShowDialog();
            string ProjectPath = fileSearch.FileName.ToString();
            if (string.IsNullOrEmpty(ProjectPath) == false)
            {
                OpenProject(ProjectPath);
            }
        }
        /// <summary>
        /// Function which is called in upper event
        /// </summary>
        /// <param name="ProjectPath"></param>
        private void OpenProject(string ProjectPath)
        {
            try
            {
                MyProject = MyTiaPortal.Projects.Open(new FileInfo(ProjectPath));
                txt_Status.Text = "Project " + ProjectPath + " opened";

            }
            catch (Exception ex)
            {
                txt_Status.Text = "Error while opening project" + ex.Message;
            }
            btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            btn_AddHW.Enabled = true;
        }
        /// <summary>
        /// SIEMENS function - event for a SAVE PROJECT button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveProject(object sender, EventArgs e)
        {
            MyProject.Save();
            txt_Status.Text = "Project saved";
        }
        /// <summary>
        /// SIEMENS function - event for a CLOSE PROJECT button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseProject(object sender, EventArgs e)
        {
            MyProject.Close();
            txt_Status.Text = "Project closed";
            btn_SearchProject.Enabled = true;
            btn_CloseProject.Enabled = false;
            btn_Save.Enabled = false;
            btn_CompileHW.Enabled = false;
        }
        /// <summary>
        /// SIEMENS function - event for a COMPILE button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Compile(object sender, EventArgs e)
        {
            btn_CompileHW.Enabled = false;
            string devname = txt_Device.Text;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = controllerTarget.GetService<ICompilable>();
                                    CompilerResult result = compiler.Compile();
                                    txt_Status.Text = "Compiling of " + controllerTarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;
                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = hmitarget.GetService<ICompilable>();
                                    CompilerResult result = compiler.Compile();
                                    txt_Status.Text = "Compiling of " + hmitarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;
                                }

                            }
                        }
                    }
                }
            }
            if (found == false)
            {
                txt_Status.Text = "Found no device with name " + txt_Device.Text;
            }

            btn_CompileHW.Enabled = true;
        }
        private void btn_AddHW_Click(object sender, EventArgs e)
        {
            AddHW();    
        }
        private void AddHW()
        {
            btn_AddHW.Enabled = false;
            string MLFB = "OrderNumber:" + txt_OrderNo.Text + "/" + txt_Version.Text;

            string name = txt_AddDevice.Text;
            string devname = "station" + txt_AddDevice.Text;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;

                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;

                                }

                            }
                        }
                    }
                }
            }
            if (found == true)
            {
                txt_Status.Text = "Device " + txt_Device.Text + " already exists";
            }
            else
            {
                Device deviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);

                txt_Status.Text = "Add Device Name: " + name + " with Order Number: " + txt_OrderNo.Text + " and Firmware Version: " + txt_Version.Text;
            }

            btn_AddHW.Enabled = true;
        }
        private void AddHW(string projectName, string typeNumber, string versionNumber)
        {
            //btn_AddHW.Enabled = false;
            string MLFB = "OrderNumber:" + typeNumber + "/" + versionNumber;

            string name = projectName;
            string devname = "station" + projectName;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;

                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;

                                }

                            }
                        }
                    }
                }
            }
            if (found == true)
            {
                txt_Status.Text = "Device " + projectName + " already exists";
            }
            else
            {
                Device deviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);

                txt_Status.Text = "Add Device Name: " + name + " with Order Number: " + typeNumber + " and Firmware Version: " + versionNumber;
            }

            //btn_AddHW.Enabled = true;
        }
        private void btn_ConnectTIA(object sender, EventArgs e)
        {
            btn_Connect.Enabled = false;
            IList<TiaPortalProcess> processes = TiaPortal.GetProcesses();
            switch (processes.Count)
            {
                case 1:
                    _tiaProcess = processes[0];
                    MyTiaPortal = _tiaProcess.Attach();
                    if(MyTiaPortal.GetCurrentProcess().Mode == TiaPortalMode.WithUserInterface)
                    {
                        rdb_WithUI.Checked = true;
                    }
                    else
                    {
                        rdb_WithoutUI.Checked = true;
                    }
                    if (MyTiaPortal.Projects.Count <= 0)
                    {
                        txt_Status.Text = "No TIA Portal Project was found!";
                        btn_Connect.Enabled = true;
                        return;
                    }
                    MyProject = MyTiaPortal.Projects[0];
                    break;
                case 0:
                    txt_Status.Text = "No running instance of TIA Portal was found!";
                    btn_Connect.Enabled = true;
                    return;
                default:
                    txt_Status.Text = "More than one running instance of TIA Portal was found!";
                    btn_Connect.Enabled = true;
                    return;
            }
            txt_Status.Text = _tiaProcess.ProjectPath.ToString();
            btn_Start.Enabled = false;
            btn_Connect.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            btn_AddHW.Enabled = true;
        }
        /// <summary>
        /// функция открытия и выбора Ексель файла
        /// добавлена переменная 
        /// </summary>
        /// <param name="UsingDragDrop"></param>
        /// <param name="e"></param>
        /// <param name="sheet">переменная которая устанавливает номер листа Ексель 
        /// по умолчанию после открытия файла устанавливается на 1 вкладку
        /// в будущем с помощью Комбобокс можно будет выбирать номер листа</param>
        private void GetObjectsData(bool UsingDragDrop, DragEventArgs e = null, int sheet = 1)
        {
            string fileName;

            if (UsingDragDrop)
            {
                string[] files;
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
                fileName = files[0].ToString();
            }
            else
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Office | *.xl*";
                ofd.ShowDialog();
                fileName = ofd.FileName;
            }

            if (fileName != string.Empty)
            {
                //MessageBox.Show(fileName);
                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(fileName);
                Worksheet ws = wb.Worksheets[sheet];
                Range ur = ws.UsedRange;
                // необходимо переделать установку наименования Датагрид таким оразом,
                // что бы названия брались с файла, а не устанавливались в ручную

                dataGridView1.Columns.Clear();
                for (int k = 1; k <= ur.Columns.Count; k++)
                {
                    dataGridView1.Columns.Add(ur.Cells[1, k].Text, ur.Cells[2, k].Text);
                }
                #region изначальное добавление наименования колонок
                //dataGridView1.Columns.Add("Column1", "Name");
                //dataGridView1.Columns.Add("Column2", "Type");
                //dataGridView1.Columns.Add("Column3", "Desc EN");
                //dataGridView1.Columns.Add("Column4", "Desc UA");
                //dataGridView1.Columns.Add("Column5", "Desc RU");
                //dataGridView1.Columns.Add("Column6", "Desc DE");
                //dataGridView1.Columns.Add("Column7", "Desc IT");
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 7].Text, ur.Cells[2, 1 + 7].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 8].Text, ur.Cells[2, 1 + 8].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 9].Text, ur.Cells[2, 1 + 9].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 10].Text, ur.Cells[2, 1 + 10].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 11].Text, ur.Cells[2, 1 + 11].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 12].Text, ur.Cells[2, 1 + 12].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 13].Text, ur.Cells[2, 1 + 13].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 14].Text, ur.Cells[2, 1 + 14].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 15].Text, ur.Cells[2, 1 + 15].Text);
                //dataGridView1.Columns.Add(ur.Cells[1, 1 + 16].Text, ur.Cells[2, 1 + 16].Text);
                #endregion
                dataGridView1.Rows.Clear();

                string [] excellRows = new string[ur.Columns.Count];
                // Создаю массив для записи значений каждой ячейки строки для дальшего добавления в ДГВ - datagridwiev
                for (int r = 3; r <= ur.Rows.Count; r++)
                {
                    for (int i = 0; i < ur.Columns.Count; i++)
                    {
                        excellRows[i] = ur.Cells[r, i + 1].Text;
                    }
                    dataGridView1.Rows.Add(excellRows);
                #region изначальное добавление строк
                    //dataGridView1.Rows.Add(ur.Cells[r, 1].Text,
                    //    ur.Cells[r, 2].Text,
                    //    ur.Cells[r, 3].Text,
                    //    ur.Cells[r, 4].Text,
                    //    ur.Cells[r, 5].Text,
                    //    ur.Cells[r, 6].Text,
                    //    ur.Cells[r, 7].Text,
                    //    ur.Cells[r, 8].Text,
                    //    ur.Cells[r, 9].Text,
                    //    ur.Cells[r, 10].Text,
                    //    ur.Cells[r, 11].Text,
                    //    ur.Cells[r, 12].Text,
                    //    ur.Cells[r, 13].Text,
                    //    ur.Cells[r, 14].Text,
                    //    ur.Cells[r, 15].Text,
                    //    ur.Cells[r, 16].Text,
                    //    ur.Cells[r, 17].Text);
                    #endregion
                }
                wb.Close();
                excel.Quit();
            }
        }
        private void btn_OpnExel_Click(object sender, EventArgs e)
        {
            GetObjectsData(false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value.ToString())||string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[1].Value.ToString())|| string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value.ToString()))
                {
                    continue;
                }
                AddHW(dataGridView1.Rows[i].Cells[0].Value.ToString(), dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString());
            }
        }
    }


}
