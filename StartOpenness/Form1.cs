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
using System.Drawing;
using System.Windows.Forms;

using _Excel = Microsoft.Office.Interop.Excel;

namespace StartOpenness
{
    public partial class Form1 : Form
    {
        private static TiaPortalProcess _tiaProcess;
        private string textMessageForRichTextBox1;
        public string TextMessageForRichTextBox1
        {
            get
            {
                return textMessageForRichTextBox1;
            }
            set
            {
                textMessageForRichTextBox1 = "[" + DateTime.Now + "] " + value+"\n";
            }
        }
        public string MyFileName
        {
            get; set;
        }
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
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
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
                richTextBox1.SelectionColor = Color.Green;
                TextMessageForRichTextBox1 = "TIA Portal started without user interface";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
                _tiaProcess = TiaPortal.GetProcesses()[0];
            }
            else
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                richTextBox1.SelectionColor = Color.Green;
                TextMessageForRichTextBox1 = "TIA Portal started with user interface";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }

            btn_SearchProject.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_Start.Enabled = false;

        }
       
        private void DisposeTIA(object sender, EventArgs e)
        {
            MyTiaPortal.Dispose();
            richTextBox1.SelectionColor = Color.Green;
            TextMessageForRichTextBox1 = "TIA Portal disposed";
            richTextBox1.SelectedText = TextMessageForRichTextBox1;
            btn_Start.Enabled = true;
            btn_Dispose.Enabled = false;
            btn_CloseProject.Enabled = false;
            btn_SearchProject.Enabled = false;
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
                richTextBox1.SelectionColor = Color.Green;
                TextMessageForRichTextBox1 = $"Project {ProjectPath} opened";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;

            }
            catch (Exception ex)
            {
                richTextBox1.SelectionColor = Color.Red;
                TextMessageForRichTextBox1 = $"Error while opening project\n{ex.Message}";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
        }
        
        private void SaveProject(object sender, EventArgs e)
        {
            MyProject.Save();
            richTextBox1.SelectionColor = Color.Green;
            TextMessageForRichTextBox1 = "Project saved";
            richTextBox1.SelectedText = TextMessageForRichTextBox1;
        }
       
        private void CloseProject(object sender, EventArgs e)
        {
            MyProject.Close();
            richTextBox1.SelectionColor = Color.Green;
            TextMessageForRichTextBox1 = "Project closed";
            richTextBox1.SelectedText = TextMessageForRichTextBox1;
            btn_SearchProject.Enabled = true;
            btn_CloseProject.Enabled = false;
            btn_Save.Enabled = false;
        }

        #endregion
        private void AddHW(string numberDeviceItemInExelFile, string deviceItemName, string deviceName, string typeNumber, string versionNumber)
        {
            string rowNumber = numberDeviceItemInExelFile;
            string MLFB = $"OrderNumber:{typeNumber}/{versionNumber}";
            string name = deviceItemName;
            string devname = deviceName;
            bool found = false;
            try
            {
                
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
                        if (deviceItem.Name == name || device.Name == devname )
                        {
                            found = true;
                        }
                    }
                }
                if (found == true)
                {
                    richTextBox1.SelectionColor = Color.Blue;
                    TextMessageForRichTextBox1 = $"DeviceItem {name} already exists";
                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                }
                else
                {
                    Device createdDeviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);
                    richTextBox1.SelectionColor = Color.Black;
                    TextMessageForRichTextBox1 = $"Added DeviceItem: {name} with {MLFB}";
                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                }
            }
            catch (Exception ex)
            {
                richTextBox1.SelectionColor = Color.Red;
                TextMessageForRichTextBox1 = $"\nRow number-{rowNumber}\nDeviceItemName-{deviceItemName}\n{ex.Message}";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }
            
            

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
                        richTextBox1.SelectionColor = Color.Red;
                        TextMessageForRichTextBox1 = "No TIA Portal Project was found!";
                        richTextBox1.SelectedText = TextMessageForRichTextBox1;
                        btn_Connect.Enabled = true;
                        return;
                    }
                    MyProject = MyTiaPortal.Projects[0];
                    break;
                case 0:
                    richTextBox1.SelectionColor = Color.Red;
                    TextMessageForRichTextBox1 = "No running instance of TIA Portal was found!";
                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                    btn_Connect.Enabled = true;
                    return;
                default:
                    richTextBox1.SelectionColor = Color.Red;
                    TextMessageForRichTextBox1 = "More than one running instance of TIA Portal was found!";
                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                    btn_Connect.Enabled = true;
                    return;
            }
            richTextBox1.SelectionColor = Color.Green;
            TextMessageForRichTextBox1 = $"Connected to project\n{_tiaProcess.ProjectPath.ToString()}";
            richTextBox1.SelectedText = TextMessageForRichTextBox1;
            btn_Start.Enabled = false;
            btn_Connect.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
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

                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(fileName);
                // Worksheet - это листы excel, выбираются из comboBox1.SelectedItem, который подгружается при загрузке файла
                // Загрузка листа меняется при изменении comboBox1.SelectedItem
                Worksheet ws = wb.Worksheets[comboBox1.SelectedItem];
                Range ur = ws.UsedRange;
              

                dataGridView1.Columns.Clear();
                // С excel все печально, нумерция ячеек начинается не с [0,0], а с [1,1]
                // Добавление колонок в ДГВ
                for (int k = 1; k <= ur.Columns.Count; k++)
                {
                    dataGridView1.Columns.Add(ur.Cells[1, k].Text, ur.Cells[1, k].Text);
                }
                dataGridView1.Rows.Clear();
                // Создаю массив для записи значений каждой ячейки строки для дальшего добавления в ДГВ
                string[] excellRows = new string[ur.Columns.Count];
                // налало начинается с 1, так что 2 строка будет 2
                for (int r = 2; r <= ur.Rows.Count; r++)
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

                AddHW(dataGridView1.Rows[i].Cells[0].Value.ToString(), dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(), dataGridView1.Rows[i].Cells[3].Value.ToString(), dataGridView1.Rows[i].Cells[4].Value.ToString());
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Subnet mySubnet = CreateNewSubnet();
            ConnectToNewSubnet(mySubnet);


            
            
        }

        private Subnet CreateNewSubnet()
        {
            Subnet MySubnet = null;
            bool foundSubnet = false;


            foreach (var subnet in MyProject.Subnets)
            {
                if (subnet.Name == dataGridView1.Rows[1].Cells[5].Value.ToString())
                {
                    foundSubnet = true;
                    MySubnet = subnet;
                }
            }
            if (foundSubnet)
            {
                richTextBox1.SelectionColor = Color.Blue;
                TextMessageForRichTextBox1 = $"Subnet with name [{dataGridView1.Rows[1].Cells[5].Value.ToString()}] alredy exist";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }
            else
            {
                MySubnet = MyProject.Subnets.Create("System:Subnet.Ethernet", dataGridView1.Rows[1].Cells[5].Value.ToString());
            }
            return MySubnet;
        }
        private void ConnectToNewSubnet(Subnet existingSubnet)
        {
            NetworkInterface network;
            Node node;
            foreach (Device device in MyProject.Devices)
            {
                foreach (DeviceItem Dev1 in device.DeviceItems)
                {
                    foreach (DeviceItem Dev2 in Dev1.DeviceItems)
                    {
                        // хардкодим название девайса для подключения, потому что при использовании 
                        //node = network.Nodes.First() - он перебирает все интерфейсы для подключения, которые нам не нужны
                        if (Dev2.Name == "PROFINET interface_1" || Dev2.Name == "PROFINET interface" || Dev2.Name == "PROFINET Interface_1" || Dev2.Name == "SCALANCE interface_1")
                        {
                            //очень полезная штука для быстрого поиска интерфейса в девайситеме
                            network = ((IEngineeringServiceProvider)Dev2).GetService<NetworkInterface>(); 
                            node = network.Nodes.First();
                            // внизу мы берем подсеть к которой подключени нод, надо для проверки
                            Subnet sub = node.ConnectedSubnet;
                            //если нод не подключен к сети
                            if (sub == null)
                            {
                                node.ConnectToSubnet(existingSubnet);
                                richTextBox1.SelectionColor = Color.Black;
                                TextMessageForRichTextBox1 = $"{Dev1.Name} is connected to [{existingSubnet.Name}]";
                                richTextBox1.SelectedText = TextMessageForRichTextBox1;
                            }
                            // если нод уже подключенк сети
                            if (sub != null)
                            {
                                // если нод подключен к сети которую мы создаем
                                if (existingSubnet.Name == sub.Name)
                                {
                                    richTextBox1.SelectionColor = Color.Blue;
                                    TextMessageForRichTextBox1 = $"{Dev1.Name} is already connected to [{existingSubnet.Name}]";
                                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                                }
                                // если нод подключен к сети которая уже была создана до нас
                                else
                                {
                                    richTextBox1.SelectionColor = Color.Purple;
                                    TextMessageForRichTextBox1 = $"{Dev1.Name} is connected to other [{sub.Name}]";
                                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                                }

                            }

                        }
                    }
                }
            }
        }
        

        private void button4_Click(object sender, EventArgs e)
        {
            Subnet mySubnet = null;
            foreach (var subnet in MyProject.Subnets)
            {
                mySubnet = subnet;
                richTextBox1.SelectionColor = Color.Black;
                TextMessageForRichTextBox1 = $"[{subnet.Name}] is founded";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }
            foreach (IoSystem ioSystem1 in mySubnet.IoSystems)
            {
                richTextBox1.SelectionColor = Color.Black;
                TextMessageForRichTextBox1 = $"[{ioSystem1.Name}] is founded";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
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
                                richTextBox1.SelectionColor = Color.Black;
                                TextMessageForRichTextBox1 = "Bingo IO Controller";
                                richTextBox1.SelectedText = TextMessageForRichTextBox1;
                                IoControllerComposition ioControllers = networkInterface.IoControllers;
                                IoController ioController = ioControllers.First();
                                if (ioController.IoSystem != null)
                                {
                                    richTextBox1.SelectionColor = Color.Blue;
                                    TextMessageForRichTextBox1 = $"{ioController.IoSystem.Name} IO system is already connected";
                                    richTextBox1.SelectedText = TextMessageForRichTextBox1;
                                }
                                if ((ioController != null)&&(ioController.IoSystem==null))
                                {
                                    ioSystem = ioController.CreateIoSystem("");
                                }
                               
                                
                            }
                            if ((networkInterface.InterfaceOperatingMode & InterfaceOperatingModes.IoDevice) != 0)
                            {
                                richTextBox1.SelectionColor = Color.Black;
                                TextMessageForRichTextBox1 = "Bingo IO Device";
                                richTextBox1.SelectedText = TextMessageForRichTextBox1;
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
                richTextBox1.SelectionColor = Color.Black;
                TextMessageForRichTextBox1 = "Bingo!";
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
            }
            else
            {
                richTextBox1.SelectionColor = Color.Black;
                TextMessageForRichTextBox1 = PLC_1.DeviceItems[0].Name;
                richTextBox1.SelectedText = TextMessageForRichTextBox1;
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
