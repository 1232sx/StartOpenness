using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Reflection;
using System.IO;

namespace StartOpenness
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_Start = new System.Windows.Forms.Button();
            this.btn_Dispose = new System.Windows.Forms.Button();
            this.btn_SearchProject = new System.Windows.Forms.Button();
            this.txt_Status = new System.Windows.Forms.TextBox();
            this.btn_CloseProject = new System.Windows.Forms.Button();
            this.btn_CompileHW = new System.Windows.Forms.Button();
            this.txt_Device = new System.Windows.Forms.TextBox();
            this.btn_Save = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.grp_TIA = new System.Windows.Forms.GroupBox();
            this.rdb_WithoutUI = new System.Windows.Forms.RadioButton();
            this.rdb_WithUI = new System.Windows.Forms.RadioButton();
            this.grp_Compile = new System.Windows.Forms.GroupBox();
            this.grp_Project = new System.Windows.Forms.GroupBox();
            this.btn_Connect = new System.Windows.Forms.Button();
            this.btn_AddHW = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_Version = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_OrderNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_AddDevice = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_OpnExel = new System.Windows.Forms.Button();
            this.btn_AddDevFrExcell = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.grp_TIA.SuspendLayout();
            this.grp_Compile.SuspendLayout();
            this.grp_Project.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(36, 85);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(115, 36);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "Start TIA";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.StartTIA);
            // 
            // btn_Dispose
            // 
            this.btn_Dispose.Enabled = false;
            this.btn_Dispose.Location = new System.Drawing.Point(36, 140);
            this.btn_Dispose.Name = "btn_Dispose";
            this.btn_Dispose.Size = new System.Drawing.Size(115, 36);
            this.btn_Dispose.TabIndex = 4;
            this.btn_Dispose.Text = "Dispose TIA";
            this.btn_Dispose.UseVisualStyleBackColor = true;
            this.btn_Dispose.Click += new System.EventHandler(this.DisposeTIA);
            // 
            // btn_SearchProject
            // 
            this.btn_SearchProject.Enabled = false;
            this.btn_SearchProject.Location = new System.Drawing.Point(38, 23);
            this.btn_SearchProject.Name = "btn_SearchProject";
            this.btn_SearchProject.Size = new System.Drawing.Size(115, 36);
            this.btn_SearchProject.TabIndex = 5;
            this.btn_SearchProject.Text = "Open Project";
            this.btn_SearchProject.UseVisualStyleBackColor = true;
            this.btn_SearchProject.Click += new System.EventHandler(this.SearchProject);
            // 
            // txt_Status
            // 
            this.txt_Status.BackColor = System.Drawing.SystemColors.Menu;
            this.txt_Status.Location = new System.Drawing.Point(12, 221);
            this.txt_Status.Name = "txt_Status";
            this.txt_Status.Size = new System.Drawing.Size(950, 20);
            this.txt_Status.TabIndex = 7;
            // 
            // btn_CloseProject
            // 
            this.btn_CloseProject.Enabled = false;
            this.btn_CloseProject.Location = new System.Drawing.Point(38, 152);
            this.btn_CloseProject.Name = "btn_CloseProject";
            this.btn_CloseProject.Size = new System.Drawing.Size(115, 36);
            this.btn_CloseProject.TabIndex = 8;
            this.btn_CloseProject.Text = "Close Project";
            this.btn_CloseProject.UseVisualStyleBackColor = true;
            this.btn_CloseProject.Click += new System.EventHandler(this.CloseProject);
            // 
            // btn_CompileHW
            // 
            this.btn_CompileHW.Enabled = false;
            this.btn_CompileHW.Location = new System.Drawing.Point(37, 85);
            this.btn_CompileHW.Name = "btn_CompileHW";
            this.btn_CompileHW.Size = new System.Drawing.Size(115, 36);
            this.btn_CompileHW.TabIndex = 12;
            this.btn_CompileHW.Text = "Compile";
            this.btn_CompileHW.UseVisualStyleBackColor = true;
            this.btn_CompileHW.Click += new System.EventHandler(this.Compile);
            // 
            // txt_Device
            // 
            this.txt_Device.Location = new System.Drawing.Point(37, 41);
            this.txt_Device.Name = "txt_Device";
            this.txt_Device.Size = new System.Drawing.Size(115, 20);
            this.txt_Device.TabIndex = 13;
            this.txt_Device.Text = "PLC_1";
            // 
            // btn_Save
            // 
            this.btn_Save.Enabled = false;
            this.btn_Save.Location = new System.Drawing.Point(38, 108);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(115, 36);
            this.btn_Save.TabIndex = 14;
            this.btn_Save.Text = "Save Project";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.SaveProject);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(37, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 15;
            this.label1.Text = "Type Device name";
            // 
            // grp_TIA
            // 
            this.grp_TIA.Controls.Add(this.rdb_WithoutUI);
            this.grp_TIA.Controls.Add(this.rdb_WithUI);
            this.grp_TIA.Controls.Add(this.btn_Dispose);
            this.grp_TIA.Controls.Add(this.btn_Start);
            this.grp_TIA.Location = new System.Drawing.Point(12, 12);
            this.grp_TIA.Name = "grp_TIA";
            this.grp_TIA.Size = new System.Drawing.Size(185, 203);
            this.grp_TIA.TabIndex = 16;
            this.grp_TIA.TabStop = false;
            this.grp_TIA.Text = "TIA Portal";
            // 
            // rdb_WithoutUI
            // 
            this.rdb_WithoutUI.AutoSize = true;
            this.rdb_WithoutUI.Location = new System.Drawing.Point(36, 51);
            this.rdb_WithoutUI.Name = "rdb_WithoutUI";
            this.rdb_WithoutUI.Size = new System.Drawing.Size(132, 17);
            this.rdb_WithoutUI.TabIndex = 2;
            this.rdb_WithoutUI.Text = "Without User Interface";
            this.rdb_WithoutUI.UseVisualStyleBackColor = true;
            // 
            // rdb_WithUI
            // 
            this.rdb_WithUI.AutoSize = true;
            this.rdb_WithUI.Checked = true;
            this.rdb_WithUI.Location = new System.Drawing.Point(36, 28);
            this.rdb_WithUI.Name = "rdb_WithUI";
            this.rdb_WithUI.Size = new System.Drawing.Size(117, 17);
            this.rdb_WithUI.TabIndex = 1;
            this.rdb_WithUI.TabStop = true;
            this.rdb_WithUI.Text = "With User Interface";
            this.rdb_WithUI.UseVisualStyleBackColor = true;
            // 
            // grp_Compile
            // 
            this.grp_Compile.Controls.Add(this.label1);
            this.grp_Compile.Controls.Add(this.txt_Device);
            this.grp_Compile.Controls.Add(this.btn_CompileHW);
            this.grp_Compile.Location = new System.Drawing.Point(586, 12);
            this.grp_Compile.Name = "grp_Compile";
            this.grp_Compile.Size = new System.Drawing.Size(185, 203);
            this.grp_Compile.TabIndex = 17;
            this.grp_Compile.TabStop = false;
            this.grp_Compile.Text = "Compile";
            // 
            // grp_Project
            // 
            this.grp_Project.Controls.Add(this.btn_Connect);
            this.grp_Project.Controls.Add(this.btn_Save);
            this.grp_Project.Controls.Add(this.btn_CloseProject);
            this.grp_Project.Controls.Add(this.btn_SearchProject);
            this.grp_Project.Location = new System.Drawing.Point(203, 12);
            this.grp_Project.Name = "grp_Project";
            this.grp_Project.Size = new System.Drawing.Size(185, 203);
            this.grp_Project.TabIndex = 18;
            this.grp_Project.TabStop = false;
            this.grp_Project.Text = "Project";
            // 
            // btn_Connect
            // 
            this.btn_Connect.Location = new System.Drawing.Point(38, 65);
            this.btn_Connect.Name = "btn_Connect";
            this.btn_Connect.Size = new System.Drawing.Size(115, 36);
            this.btn_Connect.TabIndex = 5;
            this.btn_Connect.Text = "Connect to open TIA Project";
            this.btn_Connect.UseVisualStyleBackColor = true;
            this.btn_Connect.Click += new System.EventHandler(this.btn_ConnectTIA);
            // 
            // btn_AddHW
            // 
            this.btn_AddHW.Enabled = false;
            this.btn_AddHW.Location = new System.Drawing.Point(431, 159);
            this.btn_AddHW.Name = "btn_AddHW";
            this.btn_AddHW.Size = new System.Drawing.Size(115, 36);
            this.btn_AddHW.TabIndex = 12;
            this.btn_AddHW.Text = "Add Device";
            this.btn_AddHW.UseVisualStyleBackColor = true;
            this.btn_AddHW.Click += new System.EventHandler(this.btn_AddHW_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_Version);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_OrderNo);
            this.groupBox1.Location = new System.Drawing.Point(395, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(185, 203);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(35, 105);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Type Version";
            // 
            // txt_Version
            // 
            this.txt_Version.Location = new System.Drawing.Point(35, 122);
            this.txt_Version.Name = "txt_Version";
            this.txt_Version.Size = new System.Drawing.Size(115, 20);
            this.txt_Version.TabIndex = 21;
            this.txt_Version.Text = "15.1.0.0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(36, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Type Order Nr";
            // 
            // txt_OrderNo
            // 
            this.txt_OrderNo.Location = new System.Drawing.Point(36, 82);
            this.txt_OrderNo.Name = "txt_OrderNo";
            this.txt_OrderNo.Size = new System.Drawing.Size(115, 20);
            this.txt_OrderNo.TabIndex = 19;
            this.txt_OrderNo.Text = "6AV2 123-2MA03-0AX0";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(431, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Type Device name";
            // 
            // txt_AddDevice
            // 
            this.txt_AddDevice.Location = new System.Drawing.Point(431, 54);
            this.txt_AddDevice.Name = "txt_AddDevice";
            this.txt_AddDevice.Size = new System.Drawing.Size(115, 20);
            this.txt_AddDevice.TabIndex = 13;
            this.txt_AddDevice.Text = "HMI_test";
            // 
            // groupBox2
            // 
            this.groupBox2.Location = new System.Drawing.Point(777, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(185, 203);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Group";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(98, 248);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(864, 336);
            this.dataGridView1.TabIndex = 19;
            // 
            // btn_OpnExel
            // 
            this.btn_OpnExel.Location = new System.Drawing.Point(3, 599);
            this.btn_OpnExel.Name = "btn_OpnExel";
            this.btn_OpnExel.Size = new System.Drawing.Size(76, 58);
            this.btn_OpnExel.TabIndex = 20;
            this.btn_OpnExel.Text = "Open Excel File";
            this.btn_OpnExel.UseVisualStyleBackColor = true;
            this.btn_OpnExel.Click += new System.EventHandler(this.btn_OpnExel_Click);
            // 
            // btn_AddDevFrExcell
            // 
            this.btn_AddDevFrExcell.Location = new System.Drawing.Point(88, 599);
            this.btn_AddDevFrExcell.Name = "btn_AddDevFrExcell";
            this.btn_AddDevFrExcell.Size = new System.Drawing.Size(75, 58);
            this.btn_AddDevFrExcell.TabIndex = 21;
            this.btn_AddDevFrExcell.Text = "Add Device from Excel";
            this.btn_AddDevFrExcell.UseVisualStyleBackColor = true;
            this.btn_AddDevFrExcell.Click += new System.EventHandler(this.btn_AddDevFrExcell_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(169, 599);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 58);
            this.button3.TabIndex = 22;
            this.button3.Text = "Create single subnet";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(250, 599);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 58);
            this.button4.TabIndex = 23;
            this.button4.Text = "Create subnet connected to HW";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(331, 599);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 24;
            this.button5.Text = "button5";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(412, 599);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 23);
            this.button6.TabIndex = 25;
            this.button6.Text = "button6";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(493, 599);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 23);
            this.button7.TabIndex = 26;
            this.button7.Text = "button7";
            this.button7.UseVisualStyleBackColor = true;
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(575, 599);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 23);
            this.button8.TabIndex = 27;
            this.button8.Text = "button8";
            this.button8.UseVisualStyleBackColor = true;
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(657, 599);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(75, 23);
            this.button9.TabIndex = 28;
            this.button9.Text = "button9";
            this.button9.UseVisualStyleBackColor = true;
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(739, 599);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(75, 23);
            this.button10.TabIndex = 29;
            this.button10.Text = "button10";
            this.button10.UseVisualStyleBackColor = true;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(820, 599);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 58);
            this.button11.TabIndex = 30;
            this.button11.Text = "Test";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label5.Location = new System.Drawing.Point(13, 248);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(76, 67);
            this.label5.TabIndex = 31;
            this.label5.Text = "Выпадающий список для выбора страницы Excel";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label5.UseCompatibleTextRendering = true;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(12, 343);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(77, 21);
            this.comboBox1.TabIndex = 32;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(976, 669);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button11);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btn_AddDevFrExcell);
            this.Controls.Add(this.btn_OpnExel);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txt_AddDevice);
            this.Controls.Add(this.grp_Project);
            this.Controls.Add(this.btn_AddHW);
            this.Controls.Add(this.grp_Compile);
            this.Controls.Add(this.grp_TIA);
            this.Controls.Add(this.txt_Status);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Openness_Start";
            this.grp_TIA.ResumeLayout(false);
            this.grp_TIA.PerformLayout();
            this.grp_Compile.ResumeLayout(false);
            this.grp_Compile.PerformLayout();
            this.grp_Project.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }



        #endregion

        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.Button btn_Dispose;
        private System.Windows.Forms.Button btn_SearchProject;
        private System.Windows.Forms.TextBox txt_Status;
        private System.Windows.Forms.Button btn_CloseProject;
        private System.Windows.Forms.Button btn_CompileHW;
        private System.Windows.Forms.TextBox txt_Device;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grp_TIA;
        private System.Windows.Forms.RadioButton rdb_WithoutUI;
        private System.Windows.Forms.RadioButton rdb_WithUI;
        private System.Windows.Forms.GroupBox grp_Compile;
        private System.Windows.Forms.GroupBox grp_Project;
        private System.Windows.Forms.Button btn_AddHW;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_AddDevice;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_OrderNo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_Version;
        private System.Windows.Forms.Button btn_Connect;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_OpnExel;
        private System.Windows.Forms.Button btn_AddDevFrExcell;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}

