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
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }

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

        private void SaveProject(object sender, EventArgs e)
        {
            MyProject.Save();
            txt_Status.Text = "Project saved";
        }


        private void CloseProject(object sender, EventArgs e)
        {
            MyProject.Close();

            txt_Status.Text = "Project closed";

            btn_SearchProject.Enabled = true;
            btn_CloseProject.Enabled = false;
            btn_Save.Enabled = false;
            btn_CompileHW.Enabled = false;


        }

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
    }


}
