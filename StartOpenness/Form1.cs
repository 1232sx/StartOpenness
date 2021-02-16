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
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using _Excel = Microsoft.Office.Interop.Excel;

namespace StartOpenness
{
    public partial class Form1 : Form
    {
        private static TiaPortalProcess _tiaProcess;
       
        public  Subnet MySubnet_test
        {
            get; set;
        }
        public string MyFileName
        {
            get; set;
        }
        public TiaPortal MyTiaPortal
        {
            get; set;
        }
        public string dateTimeNow => "[" + DateTime.Now + "] ";
        public Project MyProject
        {
            get; set;
        }
        public Form1()
        {
            InitializeComponent();
            //dataGridView1.AllowUserToAddRows = false;
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
            richTextBox1.Text = null;
        }
        #region Standart buttons no changes  
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

            if (filePathReg == null)
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
       
        private void StartTIA(object sender, EventArgs e)
        {
            if (rdb_WithoutUI.Checked == true)
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                richTextBox1.Text += dateTimeNow + "TIA Portal started without user interface" + System.Environment.NewLine;
                _tiaProcess = TiaPortal.GetProcesses()[0];
            }
            else
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                richTextBox1.Text += dateTimeNow +  "TIA Portal started with user interface" + System.Environment.NewLine;
            }

            btn_SearchProject.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_Start.Enabled = false;

        }
       
        private void DisposeTIA(object sender, EventArgs e)
        {
            MyTiaPortal.Dispose();
            richTextBox1.Text += dateTimeNow + "TIA Portal disposed" + System.Environment.NewLine;
            btn_Start.Enabled = true;
            btn_Dispose.Enabled = false;
            btn_CloseProject.Enabled = false;
            btn_SearchProject.Enabled = false;
            //btn_CompileHW.Enabled = false;
            btn_Save.Enabled = false;

        }
        
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
       
        private void OpenProject(string ProjectPath)
        {
            try
            {
                MyProject = MyTiaPortal.Projects.Open(new FileInfo(ProjectPath));
                richTextBox1.Text += dateTimeNow + "Project " + ProjectPath + " opened" + System.Environment.NewLine;

            }
            catch (Exception ex)
            {
                richTextBox1.Text += dateTimeNow + "Error while opening project" + ex.Message + System.Environment.NewLine;
            }
            //btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            //btn_AddHW.Enabled = true;
        }
        
        private void SaveProject(object sender, EventArgs e)
        {
            MyProject.Save();
            richTextBox1.Text += dateTimeNow + "Project saved" + System.Environment.NewLine;
        }
       
        private void CloseProject(object sender, EventArgs e)
        {
            MyProject.Close();
            richTextBox1.Text += dateTimeNow + "Project closed" + System.Environment.NewLine;
            btn_SearchProject.Enabled = true;
            btn_CloseProject.Enabled = false;
            btn_Save.Enabled = false;
            //btn_CompileHW.Enabled = false;
        }

        #endregion
      
        
        private void AddHW(string deviceItemName, string deviceName, string typeNumber, string versionNumber)
        {
            //btn_AddHW.Enabled = false;
            string MLFB = "OrderNumber:" + typeNumber + "/" + versionNumber;

            string name = deviceItemName;
            string devname = deviceName;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    // Ошибка при проверке имен 'deviceItem.Name == devname || device.Name == devname' 
                    //  device.Name == devname => device.Name == name это частично правильно
                    // добавлено еще одно условие проверки именно для HMI и в итоге получилось 
                    // if (deviceItem.Name == name || device.Name == devname|| device.Name == name)
                    // данная проверка не пропускает ни PLC ни HMI
                    if (deviceItem.Name == name || device.Name == devname || device.Name == name)
                    {
                       found = true;
                    }
                }
            }
            if (found == true)
            {
                richTextBox1.Text += dateTimeNow + "Device " + deviceItemName + " already exists" + System.Environment.NewLine;
            }
            else
            {
                Device createdDeviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);

                richTextBox1.Text += dateTimeNow + "Add Device Name: " + name + " with Order Number: " + typeNumber + " and Firmware Version: " + versionNumber + System.Environment.NewLine;
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
                    if (MyTiaPortal.GetCurrentProcess().Mode == TiaPortalMode.WithUserInterface)
                    {
                        rdb_WithUI.Checked = true;
                    }
                    else
                    {
                        rdb_WithoutUI.Checked = true;
                    }
                    if (MyTiaPortal.Projects.Count <= 0)
                    {
                        richTextBox1.Text += dateTimeNow + "No TIA Portal Project was found!" + System.Environment.NewLine;
                        btn_Connect.Enabled = true;
                        return;
                    }
                    MyProject = MyTiaPortal.Projects[0];
                    break;
                case 0:
                    richTextBox1.Text += dateTimeNow + "No running instance of TIA Portal was found!" + System.Environment.NewLine;
                    btn_Connect.Enabled = true;
                    return;
                default:
                    richTextBox1.Text += dateTimeNow + "More than one running instance of TIA Portal was found!" + System.Environment.NewLine;
                    btn_Connect.Enabled = true;
                    return;
            }
            richTextBox1.Text += dateTimeNow + "Connected to project " +_tiaProcess.ProjectPath.ToString() + System.Environment.NewLine;
            btn_Start.Enabled = false;
            btn_Connect.Enabled = true;
            btn_Dispose.Enabled = true;
            //btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            //btn_AddHW.Enabled = true;
        }
        private void ItinializeCombobox1(string fileName)
        {
            if (fileName != string.Empty)
            {
                //MessageBox.Show(fileName);
                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(fileName);
                // Провубю вставить выбор по номеру в Комбобоксе
                List<string> sheetName = new List<string>();
                foreach (Worksheet item in wb.Worksheets)
                {
                    sheetName.Add(item.Name);
                }
                comboBox1.DataSource = sheetName;
                wb.Close();
                excel.Quit();
            }
        }
        private void GetObjectsData(string fileName)
        {
            
            if (fileName != string.Empty)
            {

                //MessageBox.Show(fileName);
                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(fileName);
                // Провубю вставить выбор по номеру в Комбобоксе
                Worksheet ws = wb.Worksheets[comboBox1.SelectedItem];
                Range ur = ws.UsedRange;
                
                // необходимо переделать установку наименования Датагрид таким оразом,
                // что бы названия брались с файла, а не устанавливались в ручную


                dataGridView1.Columns.Clear();
                for (int k = 1; k <= ur.Columns.Count; k++)
                {
                    dataGridView1.Columns.Add(ur.Cells[1, k].Text, ur.Cells[2, k].Text);
                }
                
                dataGridView1.Rows.Clear();

                string[] excellRows = new string[ur.Columns.Count];
                // Создаю массив для записи значений каждой ячейки строки для дальшего добавления в ДГВ - datagridwiev
                for (int r = 3; r <= ur.Rows.Count; r++)
                {
                    for (int i = 0; i < ur.Columns.Count; i++)
                    {
                        excellRows[i] = ur.Cells[r, i + 1].Text;
                    }
                    dataGridView1.Rows.Add(excellRows);
                   
                }
                wb.Close();
                excel.Quit();
            }
        }
        private string FileDialogOpen(bool UsingDragDrop, DragEventArgs e = null)
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
            return fileName;
        }
        private void btn_OpnExel_Click(object sender, EventArgs e)
        {
            MyFileName = FileDialogOpen(false);
            ItinializeCombobox1(MyFileName);
           
        }
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value == null)
                {
                    continue;
                }

                AddHW(dataGridView1.Rows[i].Cells[0].Value.ToString(), dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(), dataGridView1.Rows[i].Cells[3].Value.ToString());
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Subnet MySubnet = null;
            bool foundSubnet = false;

            
            foreach (var subnet in MyProject.Subnets)
            {
                if (subnet.Name== dataGridView1.Rows[1].Cells[4].Value.ToString())
                {
                        foundSubnet = true;
                        MySubnet = subnet;
                }
            }
            if (foundSubnet)
            {
                richTextBox1.Text += dateTimeNow + "Subnet with name [" + dataGridView1.Rows[1].Cells[4].Value.ToString() + "] alredy exist" + System.Environment.NewLine;
            }
            else
            {
                MySubnet = MyProject.Subnets.Create("System:Subnet.Ethernet", dataGridView1.Rows[1].Cells[4].Value.ToString());
            }
            NetworkInterface network = null; ;
            Node node;
            
            
            int counter_device = 0;
            int counter_Dev1 = 0;
            int counter_Dev2 = 0;
            

            foreach (Device device in MyProject.Devices)
            {
                foreach (DeviceItem Dev1 in device.DeviceItems)
                {
                    foreach (DeviceItem Dev2 in Dev1.DeviceItems)
                    {
                        if (Dev2.Name == "PROFINET interface_1" || Dev2.Name == "PROFINET interface" || Dev2.Name == "PROFINET Interface_1" || Dev2.Name == "SCALANCE interface_1")
                        {
                            network = MyProject.Devices[counter_device].DeviceItems[counter_Dev1].DeviceItems[counter_Dev2].GetService<NetworkInterface>();
                            node = network.Nodes[0];
                            // внизу мы берем подсеть к которой подключени нод, надо для проверки
                            Subnet sub = node.ConnectedSubnet;
                            // если интерфейс подключеня в наличии и еще не подключен к сети
                            if (node != null&&sub==null)
                            {
                                node.ConnectToSubnet(MySubnet);
                                richTextBox1.Text += dateTimeNow + MyProject.Devices[counter_device].Name + " is connected to [" + MySubnet.Name + "]" + System.Environment.NewLine;
                            }
                            // если интерфейс подключения в наличии и уже есть подключение к сети
                            if (node!=null&&sub!=null)
                            {
                                // Проверка соответствия имен заданной сети и сети к которой подключен нод
                                if (MySubnet.Name==sub.Name)
                                {
                                    richTextBox1.Text += dateTimeNow + MyProject.Devices[counter_device].Name + " is already connected to [" + MySubnet.Name + "]" + System.Environment.NewLine;
                                }
                                else
                                {
                                    node.ConnectToSubnet(MySubnet);
                                    richTextBox1.Text += dateTimeNow + MyProject.Devices[counter_device].Name + " is connected to [" + MySubnet.Name + "]" + System.Environment.NewLine;
                                }
                            }
                           node = null;
                            
                        }
                        counter_Dev2++;

                    }
                    counter_Dev1++;
                    counter_Dev2 = 0;
                }
                counter_device++;
                counter_Dev1 = 0;
            }
            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Subnet mySubnet = null;
            foreach (var subnet in MyProject.Subnets)
            {
                mySubnet = subnet;
                richTextBox1.Text += dateTimeNow + "[" + subnet.Name + "] is founded" + System.Environment.NewLine;
            }
            foreach (IoSystem ioSystem1 in mySubnet.IoSystems)
            {
                richTextBox1.Text += dateTimeNow + "[" + ioSystem1.Name + "] is founded" + System.Environment.NewLine;
            }

            
          


            int counter_device = 0;
            int counter_Dev1 = 0;
            int counter_Dev2 = 0;
            NetworkInterface networkInterface = null;
            IoSystem ioSystem = null;
            foreach (Device device in MyProject.Devices)
            {
                foreach (DeviceItem Dev1 in device.DeviceItems)
                {
                    foreach (DeviceItem Dev2 in Dev1.DeviceItems)
                    {
                        if (Dev2.Name == "PROFINET interface_1" || Dev2.Name == "PROFINET interface" || Dev2.Name == "PROFINET Interface_1" || Dev2.Name == "SCALANCE interface_1")
                        {
                            networkInterface = MyProject.Devices[counter_device].DeviceItems[counter_Dev1].DeviceItems[counter_Dev2].GetService<NetworkInterface>();
                            if ((networkInterface.InterfaceOperatingMode & InterfaceOperatingModes.IoController) != 0)
                            {
                                richTextBox1.Text += dateTimeNow + " Bingo IO Controller" + Environment.NewLine;
                                IoControllerComposition ioControllers = networkInterface.IoControllers;
                                IoController ioController = ioControllers.First();
                                if (ioController.IoSystem != null)
                                {
                                    richTextBox1.Text += dateTimeNow + ioController.IoSystem.Name  + " IO system is already connected" + Environment.NewLine;
                                }
                                if ((ioController != null)&&(ioController.IoSystem==null))
                                {
                                    ioSystem = ioController.CreateIoSystem("");
                                }
                               
                                
                            }
                            if ((networkInterface.InterfaceOperatingMode & InterfaceOperatingModes.IoDevice) != 0)
                            {
                                richTextBox1.Text += dateTimeNow + " Bingo IO Device" + Environment.NewLine;
                                IoConnectorComposition ioConnectors = networkInterface.IoConnectors;
                                IoConnector ioConnector = ioConnectors.First();

                                if (ioConnector != null)
                                {
                                    ioConnector.ConnectToIoSystem(ioSystem);
                                }
                            }



                        }
                        counter_Dev2++;

                    }
                    counter_Dev1++;
                    counter_Dev2 = 0;
                }
                counter_device++;
                counter_Dev1 = 0;
            }




        }
        private void button11_Click(object sender, EventArgs e)
        {
          


             
        }
        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("Device", "Device");
            dataGridView1.Columns.Add("Dev1", "Dev1");
            dataGridView1.Columns.Add("Dev2", "Dev2");
            dataGridView1.Columns.Add("Dev3", "Dev3");
            dataGridView1.Columns.Add("Dev4", "Dev4");
            foreach (Device device in MyProject.Devices)
            {
                dataGridView1.Rows.Add(device.Name);
                foreach (DeviceItem Dev1 in device.DeviceItems)
                {
                    dataGridView1.Rows.Add("-", Dev1.Name);
                    foreach (DeviceItem Dev2 in Dev1.DeviceItems)
                    {
                        dataGridView1.Rows.Add("-", "-", Dev2.Name);
                        foreach (DeviceItem Dev3 in Dev2.DeviceItems)
                        {
                            dataGridView1.Rows.Add("-", "-", "-", Dev3.Name);


                        }
                    }
                }
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            Device PLC_1 = MyProject.Devices.CreateWithItem("OrderNumber:6ES7 517-3FP00-0AB0/V2.1", "PLC_1", "PLC_1_station");
            
            //DeviceItemComposition deviceItems = PLC_1.DeviceItems;
            HardwareObject hwObject = PLC_1.DeviceItems[0];

            if (hwObject.CanPlugNew("OrderNumber:6ES7 521-1BH10-0AA0/V1.0", "DI 16x24VDC BA_1", 3))
            {
                DeviceItem newPluggedDeviceItem = hwObject.PlugNew("OrderNumber:6ES7 521-1BH10-0AA0/V1.0", "DI 16x24VDC BA_1", 3);
                richTextBox1.Text+= dateTimeNow + "Bingo!" + System.Environment.NewLine;
            }
            else
            {
                richTextBox1.Text += dateTimeNow + PLC_1.DeviceItems[0].Name + System.Environment.NewLine;
            }

        }
        private void button7_Click(object sender, EventArgs e)
        {
           
            

            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            GetObjectsData(MyFileName);
        }

        
    }

}
