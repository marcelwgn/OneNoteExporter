//OneNoteExporter: export sections from OneNote to Word
//Copyright(C) 2017 Marcel Wagner
//This program is free software; you can redistribute it and/or modify it under the terms 
//of the GNU General Public License as published by the Free Software Foundation; either 
//version 3 of the License, or(at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
//without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
//See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program; 
//if not, see<http://www.gnu.org/licenses/>. 

using System;
using System.Windows.Forms;

namespace OneNoteExporter
{
    partial class Mainframe
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mainframe));
            this.sectionLoadButton = new System.Windows.Forms.Button();
            this.dgvSection = new System.Windows.Forms.DataGridView();
            this.doWork = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.SectionGroup = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Section = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TimeExported = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.exportingProgressbar = new System.Windows.Forms.ProgressBar();
            this.progressLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.sectionNameBox = new System.Windows.Forms.MaskedTextBox();
            this.checkAllSections = new System.Windows.Forms.Button();
            this.uncheckAllButton = new System.Windows.Forms.Button();
            this.applyPresetsButton = new System.Windows.Forms.Button();
            this.saveExports = new System.Windows.Forms.Button();
            this.exportsClosingCheck = new System.Windows.Forms.CheckBox();
            this.exportFolderLabel = new System.Windows.Forms.Label();
            this.exportFolderBox = new System.Windows.Forms.MaskedTextBox();
            this.dgvNotebook = new System.Windows.Forms.DataGridView();
            this.notebookColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSection)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNotebook)).BeginInit();
            this.SuspendLayout();
            // 
            // sectionLoadButton
            // 
            this.sectionLoadButton.Location = new System.Drawing.Point(15, 10);
            this.sectionLoadButton.Name = "sectionLoadButton";
            this.sectionLoadButton.Size = new System.Drawing.Size(177, 23);
            this.sectionLoadButton.TabIndex = 0;
            this.sectionLoadButton.Text = "Load sections";
            this.sectionLoadButton.UseVisualStyleBackColor = true;
            this.sectionLoadButton.Click += new System.EventHandler(this.loadSections);
            // 
            // dgvSection
            // 
            this.dgvSection.AllowUserToAddRows = false;
            this.dgvSection.AllowUserToDeleteRows = false;
            this.dgvSection.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSection.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvSection.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSection.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.doWork,
            this.SectionGroup,
            this.Section,
            this.fileName,
            this.TimeExported});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSection.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvSection.Location = new System.Drawing.Point(226, 69);
            this.dgvSection.Name = "dgvSection";
            this.dgvSection.RowHeadersVisible = false;
            this.dgvSection.ShowEditingIcon = false;
            this.dgvSection.Size = new System.Drawing.Size(1303, 425);
            this.dgvSection.TabIndex = 4;
            this.dgvSection.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_cellClick);
            this.dgvSection.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSection_CellValueChanged);
            // 
            // doWork
            // 
            this.doWork.FalseValue = "0";
            this.doWork.HeaderText = "Export";
            this.doWork.Name = "doWork";
            this.doWork.TrueValue = "1";
            // 
            // SectionGroup
            // 
            this.SectionGroup.HeaderText = "Section group";
            this.SectionGroup.Name = "SectionGroup";
            this.SectionGroup.ReadOnly = true;
            this.SectionGroup.Width = 300;
            // 
            // Section
            // 
            this.Section.HeaderText = "Section";
            this.Section.Name = "Section";
            this.Section.Width = 300;
            // 
            // fileName
            // 
            this.fileName.HeaderText = "Export name";
            this.fileName.Name = "fileName";
            this.fileName.Width = 300;
            // 
            // TimeExported
            // 
            this.TimeExported.HeaderText = "Time exported";
            this.TimeExported.Name = "TimeExported";
            this.TimeExported.Width = 300;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(246, 10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(177, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Export sections to docx";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.export_clicked);
            // 
            // exportingProgressbar
            // 
            this.exportingProgressbar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.exportingProgressbar.Location = new System.Drawing.Point(15, 500);
            this.exportingProgressbar.Name = "exportingProgressbar";
            this.exportingProgressbar.Size = new System.Drawing.Size(1513, 23);
            this.exportingProgressbar.TabIndex = 8;
            // 
            // progressLabel
            // 
            this.progressLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(12, 526);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(38, 13);
            this.progressLabel.TabIndex = 9;
            this.progressLabel.Text = "Ready";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Section name:";
            // 
            // sectionNameBox
            // 
            this.sectionNameBox.Location = new System.Drawing.Point(92, 43);
            this.sectionNameBox.Name = "sectionNameBox";
            this.sectionNameBox.Size = new System.Drawing.Size(100, 20);
            this.sectionNameBox.TabIndex = 11;
            // 
            // checkAllSections
            // 
            this.checkAllSections.Location = new System.Drawing.Point(760, 10);
            this.checkAllSections.Name = "checkAllSections";
            this.checkAllSections.Size = new System.Drawing.Size(111, 23);
            this.checkAllSections.TabIndex = 12;
            this.checkAllSections.Text = "Check all sections";
            this.checkAllSections.UseVisualStyleBackColor = true;
            this.checkAllSections.Click += new System.EventHandler(this.checkAllSections_Click);
            // 
            // uncheckAllButton
            // 
            this.uncheckAllButton.Location = new System.Drawing.Point(760, 39);
            this.uncheckAllButton.Name = "uncheckAllButton";
            this.uncheckAllButton.Size = new System.Drawing.Size(142, 23);
            this.uncheckAllButton.TabIndex = 13;
            this.uncheckAllButton.Text = "Uncheck all sections";
            this.uncheckAllButton.UseVisualStyleBackColor = true;
            this.uncheckAllButton.Click += new System.EventHandler(this.uncheckAllButton_Click);
            // 
            // applyPresetsButton
            // 
            this.applyPresetsButton.Location = new System.Drawing.Point(908, 39);
            this.applyPresetsButton.Name = "applyPresetsButton";
            this.applyPresetsButton.Size = new System.Drawing.Size(95, 23);
            this.applyPresetsButton.TabIndex = 14;
            this.applyPresetsButton.Text = "Apply settings";
            this.applyPresetsButton.UseVisualStyleBackColor = true;
            this.applyPresetsButton.Click += new System.EventHandler(this.applyPresets_Click);
            // 
            // saveExports
            // 
            this.saveExports.Location = new System.Drawing.Point(1175, 11);
            this.saveExports.Name = "saveExports";
            this.saveExports.Size = new System.Drawing.Size(177, 23);
            this.saveExports.TabIndex = 15;
            this.saveExports.Text = "Save settings";
            this.saveExports.UseVisualStyleBackColor = true;
            this.saveExports.Click += new System.EventHandler(this.saveSettings);
            // 
            // exportsClosingCheck
            // 
            this.exportsClosingCheck.AutoSize = true;
            this.exportsClosingCheck.Location = new System.Drawing.Point(1175, 43);
            this.exportsClosingCheck.Name = "exportsClosingCheck";
            this.exportsClosingCheck.Size = new System.Drawing.Size(155, 17);
            this.exportsClosingCheck.TabIndex = 16;
            this.exportsClosingCheck.Text = "Save settings when closing";
            this.exportsClosingCheck.UseVisualStyleBackColor = true;
            // 
            // exportFolderLabel
            // 
            this.exportFolderLabel.AutoSize = true;
            this.exportFolderLabel.Location = new System.Drawing.Point(461, 15);
            this.exportFolderLabel.Name = "exportFolderLabel";
            this.exportFolderLabel.Size = new System.Drawing.Size(241, 13);
            this.exportFolderLabel.TabIndex = 17;
            this.exportFolderLabel.Text = "Exportfolder (Subfolder of location of this program)";
            // 
            // exportFolderBox
            // 
            this.exportFolderBox.Location = new System.Drawing.Point(464, 42);
            this.exportFolderBox.Name = "exportFolderBox";
            this.exportFolderBox.Size = new System.Drawing.Size(238, 20);
            this.exportFolderBox.TabIndex = 18;
            this.exportFolderBox.TextChanged += new System.EventHandler(this.exportFolderBox_TextChanged);
            // 
            // dgvNotebook
            // 
            this.dgvNotebook.AllowUserToAddRows = false;
            this.dgvNotebook.AllowUserToDeleteRows = false;
            this.dgvNotebook.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvNotebook.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.notebookColumn});
            this.dgvNotebook.Location = new System.Drawing.Point(15, 69);
            this.dgvNotebook.Name = "dgvNotebook";
            this.dgvNotebook.RowHeadersVisible = false;
            this.dgvNotebook.ShowEditingIcon = false;
            this.dgvNotebook.Size = new System.Drawing.Size(205, 425);
            this.dgvNotebook.TabIndex = 19;
            this.dgvNotebook.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvNotebook_CellClick);
            // 
            // notebookColumn
            // 
            this.notebookColumn.HeaderText = "Notebook";
            this.notebookColumn.Name = "notebookColumn";
            this.notebookColumn.ReadOnly = true;
            this.notebookColumn.Width = 202;
            // 
            // Mainframe
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1541, 548);
            this.Controls.Add(this.dgvNotebook);
            this.Controls.Add(this.exportFolderBox);
            this.Controls.Add(this.exportFolderLabel);
            this.Controls.Add(this.exportsClosingCheck);
            this.Controls.Add(this.saveExports);
            this.Controls.Add(this.applyPresetsButton);
            this.Controls.Add(this.uncheckAllButton);
            this.Controls.Add(this.checkAllSections);
            this.Controls.Add(this.sectionNameBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.exportingProgressbar);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dgvSection);
            this.Controls.Add(this.sectionLoadButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Mainframe";
            this.Text = "OneNote Section Exporter";
            ((System.ComponentModel.ISupportInitialize)(this.dgvSection)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNotebook)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }






        #endregion

        private System.Windows.Forms.Button sectionLoadButton;
        private Button button1;
        private ProgressBar exportingProgressbar;
        private Label progressLabel;
        private Label label1;
        private MaskedTextBox sectionNameBox;
        private Button checkAllSections;
        private Button uncheckAllButton;
        private Button applyPresetsButton;
        private Button saveExports;
        private CheckBox exportsClosingCheck;
        private Label exportFolderLabel;
        private MaskedTextBox exportFolderBox;
        public DataGridView dgvSection;
        private DataGridViewCheckBoxColumn doWork;
        private DataGridViewTextBoxColumn SectionGroup;
        private DataGridViewTextBoxColumn Section;
        private DataGridViewTextBoxColumn fileName;
        private DataGridViewTextBoxColumn TimeExported;
        private DataGridView dgvNotebook;
        private DataGridViewTextBoxColumn notebookColumn;
    }
}

