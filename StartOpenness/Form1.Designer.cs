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
            this.btn_AddHW = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_AddDevice = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_OrderNo = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_Version = new System.Windows.Forms.TextBox();
            this.btn_Connect = new System.Windows.Forms.Button();
            this.grp_TIA.SuspendLayout();
            this.grp_Compile.SuspendLayout();
            this.grp_Project.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(48, 105);
            this.btn_Start.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(153, 44);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "Start TIA";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.StartTIA);
            // 
            // btn_Dispose
            // 
            this.btn_Dispose.Enabled = false;
            this.btn_Dispose.Location = new System.Drawing.Point(48, 172);
            this.btn_Dispose.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Dispose.Name = "btn_Dispose";
            this.btn_Dispose.Size = new System.Drawing.Size(153, 44);
            this.btn_Dispose.TabIndex = 4;
            this.btn_Dispose.Text = "Dispose TIA";
            this.btn_Dispose.UseVisualStyleBackColor = true;
            this.btn_Dispose.Click += new System.EventHandler(this.DisposeTIA);
            // 
            // btn_SearchProject
            // 
            this.btn_SearchProject.Enabled = false;
            this.btn_SearchProject.Location = new System.Drawing.Point(51, 28);
            this.btn_SearchProject.Margin = new System.Windows.Forms.Padding(4);
            this.btn_SearchProject.Name = "btn_SearchProject";
            this.btn_SearchProject.Size = new System.Drawing.Size(153, 44);
            this.btn_SearchProject.TabIndex = 5;
            this.btn_SearchProject.Text = "Open Project";
            this.btn_SearchProject.UseVisualStyleBackColor = true;
            this.btn_SearchProject.Click += new System.EventHandler(this.SearchProject);
            // 
            // txt_Status
            // 
            this.txt_Status.BackColor = System.Drawing.SystemColors.Menu;
            this.txt_Status.Location = new System.Drawing.Point(16, 298);
            this.txt_Status.Margin = new System.Windows.Forms.Padding(4);
            this.txt_Status.Name = "txt_Status";
            this.txt_Status.Size = new System.Drawing.Size(1013, 22);
            this.txt_Status.TabIndex = 7;
            // 
            // btn_CloseProject
            // 
            this.btn_CloseProject.Enabled = false;
            this.btn_CloseProject.Location = new System.Drawing.Point(51, 187);
            this.btn_CloseProject.Margin = new System.Windows.Forms.Padding(4);
            this.btn_CloseProject.Name = "btn_CloseProject";
            this.btn_CloseProject.Size = new System.Drawing.Size(153, 44);
            this.btn_CloseProject.TabIndex = 8;
            this.btn_CloseProject.Text = "Close Project";
            this.btn_CloseProject.UseVisualStyleBackColor = true;
            this.btn_CloseProject.Click += new System.EventHandler(this.CloseProject);
            // 
            // btn_CompileHW
            // 
            this.btn_CompileHW.Enabled = false;
            this.btn_CompileHW.Location = new System.Drawing.Point(49, 105);
            this.btn_CompileHW.Margin = new System.Windows.Forms.Padding(4);
            this.btn_CompileHW.Name = "btn_CompileHW";
            this.btn_CompileHW.Size = new System.Drawing.Size(153, 44);
            this.btn_CompileHW.TabIndex = 12;
            this.btn_CompileHW.Text = "Compile";
            this.btn_CompileHW.UseVisualStyleBackColor = true;
            this.btn_CompileHW.Click += new System.EventHandler(this.Compile);
            // 
            // txt_Device
            // 
            this.txt_Device.Location = new System.Drawing.Point(49, 51);
            this.txt_Device.Margin = new System.Windows.Forms.Padding(4);
            this.txt_Device.Name = "txt_Device";
            this.txt_Device.Size = new System.Drawing.Size(152, 22);
            this.txt_Device.TabIndex = 13;
            this.txt_Device.Text = "PLC_1";
            // 
            // btn_Save
            // 
            this.btn_Save.Enabled = false;
            this.btn_Save.Location = new System.Drawing.Point(51, 133);
            this.btn_Save.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(153, 44);
            this.btn_Save.TabIndex = 14;
            this.btn_Save.Text = "Save Project";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.SaveProject);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(49, 30);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 17);
            this.label1.TabIndex = 15;
            this.label1.Text = "Type Device name";
            // 
            // grp_TIA
            // 
            this.grp_TIA.Controls.Add(this.rdb_WithoutUI);
            this.grp_TIA.Controls.Add(this.rdb_WithUI);
            this.grp_TIA.Controls.Add(this.btn_Dispose);
            this.grp_TIA.Controls.Add(this.btn_Start);
            this.grp_TIA.Location = new System.Drawing.Point(16, 15);
            this.grp_TIA.Margin = new System.Windows.Forms.Padding(4);
            this.grp_TIA.Name = "grp_TIA";
            this.grp_TIA.Padding = new System.Windows.Forms.Padding(4);
            this.grp_TIA.Size = new System.Drawing.Size(247, 250);
            this.grp_TIA.TabIndex = 16;
            this.grp_TIA.TabStop = false;
            this.grp_TIA.Text = "TIA Portal";
            // 
            // rdb_WithoutUI
            // 
            this.rdb_WithoutUI.AutoSize = true;
            this.rdb_WithoutUI.Location = new System.Drawing.Point(48, 63);
            this.rdb_WithoutUI.Margin = new System.Windows.Forms.Padding(4);
            this.rdb_WithoutUI.Name = "rdb_WithoutUI";
            this.rdb_WithoutUI.Size = new System.Drawing.Size(170, 21);
            this.rdb_WithoutUI.TabIndex = 2;
            this.rdb_WithoutUI.Text = "Without User Interface";
            this.rdb_WithoutUI.UseVisualStyleBackColor = true;
            // 
            // rdb_WithUI
            // 
            this.rdb_WithUI.AutoSize = true;
            this.rdb_WithUI.Checked = true;
            this.rdb_WithUI.Location = new System.Drawing.Point(48, 34);
            this.rdb_WithUI.Margin = new System.Windows.Forms.Padding(4);
            this.rdb_WithUI.Name = "rdb_WithUI";
            this.rdb_WithUI.Size = new System.Drawing.Size(150, 21);
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
            this.grp_Compile.Location = new System.Drawing.Point(782, 15);
            this.grp_Compile.Margin = new System.Windows.Forms.Padding(4);
            this.grp_Compile.Name = "grp_Compile";
            this.grp_Compile.Padding = new System.Windows.Forms.Padding(4);
            this.grp_Compile.Size = new System.Drawing.Size(247, 250);
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
            this.grp_Project.Location = new System.Drawing.Point(271, 15);
            this.grp_Project.Margin = new System.Windows.Forms.Padding(4);
            this.grp_Project.Name = "grp_Project";
            this.grp_Project.Padding = new System.Windows.Forms.Padding(4);
            this.grp_Project.Size = new System.Drawing.Size(247, 250);
            this.grp_Project.TabIndex = 18;
            this.grp_Project.TabStop = false;
            this.grp_Project.Text = "Project";
            // 
            // btn_AddHW
            // 
            this.btn_AddHW.Enabled = false;
            this.btn_AddHW.Location = new System.Drawing.Point(575, 196);
            this.btn_AddHW.Margin = new System.Windows.Forms.Padding(4);
            this.btn_AddHW.Name = "btn_AddHW";
            this.btn_AddHW.Size = new System.Drawing.Size(153, 44);
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
            this.groupBox1.Location = new System.Drawing.Point(527, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(247, 250);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(575, 45);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(126, 17);
            this.label2.TabIndex = 15;
            this.label2.Text = "Type Device name";
            // 
            // txt_AddDevice
            // 
            this.txt_AddDevice.Location = new System.Drawing.Point(575, 66);
            this.txt_AddDevice.Margin = new System.Windows.Forms.Padding(4);
            this.txt_AddDevice.Name = "txt_AddDevice";
            this.txt_AddDevice.Size = new System.Drawing.Size(152, 22);
            this.txt_AddDevice.TabIndex = 13;
            this.txt_AddDevice.Text = "PLC_1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(48, 80);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 17);
            this.label3.TabIndex = 20;
            this.label3.Text = "Type Order Nr";
            // 
            // txt_OrderNo
            // 
            this.txt_OrderNo.Location = new System.Drawing.Point(48, 101);
            this.txt_OrderNo.Margin = new System.Windows.Forms.Padding(4);
            this.txt_OrderNo.Name = "txt_OrderNo";
            this.txt_OrderNo.Size = new System.Drawing.Size(152, 22);
            this.txt_OrderNo.TabIndex = 19;
            this.txt_OrderNo.Text = "6ES7 516-3AN01-0AB0";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(47, 129);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 17);
            this.label4.TabIndex = 22;
            this.label4.Text = "Type Version";
            // 
            // txt_Version
            // 
            this.txt_Version.Location = new System.Drawing.Point(47, 150);
            this.txt_Version.Margin = new System.Windows.Forms.Padding(4);
            this.txt_Version.Name = "txt_Version";
            this.txt_Version.Size = new System.Drawing.Size(152, 22);
            this.txt_Version.TabIndex = 21;
            this.txt_Version.Text = "V2.1";
            // 
            // btn_Connect
            // 
            this.btn_Connect.Location = new System.Drawing.Point(51, 80);
            this.btn_Connect.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Connect.Name = "btn_Connect";
            this.btn_Connect.Size = new System.Drawing.Size(153, 44);
            this.btn_Connect.TabIndex = 5;
            this.btn_Connect.Text = "Connect to open TIA Project";
            this.btn_Connect.UseVisualStyleBackColor = true;
            this.btn_Connect.Click += new System.EventHandler(this.btn_ConnectTIA);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(1046, 337);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txt_AddDevice);
            this.Controls.Add(this.grp_Project);
            this.Controls.Add(this.btn_AddHW);
            this.Controls.Add(this.grp_Compile);
            this.Controls.Add(this.grp_TIA);
            this.Controls.Add(this.txt_Status);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Openness_Start";
            this.grp_TIA.ResumeLayout(false);
            this.grp_TIA.PerformLayout();
            this.grp_Compile.ResumeLayout(false);
            this.grp_Compile.PerformLayout();
            this.grp_Project.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
    }
}

