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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btn_Start = new System.Windows.Forms.Button();
            this.btn_Dispose = new System.Windows.Forms.Button();
            this.btn_SearchProject = new System.Windows.Forms.Button();
            this.btn_CloseProject = new System.Windows.Forms.Button();
            this.btn_Save = new System.Windows.Forms.Button();
            this.grp_TIA = new System.Windows.Forms.GroupBox();
            this.rdb_WithoutUI = new System.Windows.Forms.RadioButton();
            this.rdb_WithUI = new System.Windows.Forms.RadioButton();
            this.grp_Project = new System.Windows.Forms.GroupBox();
            this.btn_Connect = new System.Windows.Forms.Button();
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
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.grp_TIA.SuspendLayout();
            this.grp_Project.SuspendLayout();
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
            this.btn_AddDevFrExcell.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(169, 599);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 58);
            this.button3.TabIndex = 22;
            this.button3.Text = "Create subnet connect HW";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(250, 599);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 58);
            this.button4.TabIndex = 23;
            this.button4.Text = "IOsubnet";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(331, 599);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 58);
            this.button5.TabIndex = 24;
            this.button5.Text = "all devices";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(412, 599);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 58);
            this.button6.TabIndex = 25;
            this.button6.Text = "button6";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(493, 599);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 58);
            this.button7.TabIndex = 26;
            this.button7.Text = "TestIP";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(575, 599);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 58);
            this.button8.TabIndex = 27;
            this.button8.Text = "Add device item to PLC";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(657, 599);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(75, 58);
            this.button9.TabIndex = 28;
            this.button9.Text = "button9";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(739, 599);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(75, 58);
            this.button10.TabIndex = 29;
            this.button10.Text = "button10";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(820, 599);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 58);
            this.button11.TabIndex = 30;
            this.button11.Text = "Device List";
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
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(395, 13);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(567, 202);
            this.richTextBox1.TabIndex = 33;
            this.richTextBox1.Text = "";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excell|*.xls*";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(976, 669);
            this.Controls.Add(this.richTextBox1);
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
            this.Controls.Add(this.grp_Project);
            this.Controls.Add(this.grp_TIA);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Openness_Start";
            this.grp_TIA.ResumeLayout(false);
            this.grp_TIA.PerformLayout();
            this.grp_Project.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }



        #endregion

        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.Button btn_Dispose;
        private System.Windows.Forms.Button btn_SearchProject;
        private System.Windows.Forms.Button btn_CloseProject;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.GroupBox grp_TIA;
        private System.Windows.Forms.RadioButton rdb_WithoutUI;
        private System.Windows.Forms.RadioButton rdb_WithUI;
        private System.Windows.Forms.GroupBox grp_Project;
        private System.Windows.Forms.Button btn_Connect;
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
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

