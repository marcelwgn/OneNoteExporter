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
            this.sectionLoadButton = new System.Windows.Forms.Button();
            this.notebookName = new System.Windows.Forms.MaskedTextBox();
            this.notebookFilterLabel = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.Nodebook = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Section = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TimeExported = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.doWork = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.exportingProgressbar = new System.Windows.Forms.ProgressBar();
            this.progressLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // sectionLoadButton
            // 
            this.sectionLoadButton.Location = new System.Drawing.Point(13, 13);
            this.sectionLoadButton.Name = "sectionLoadButton";
            this.sectionLoadButton.Size = new System.Drawing.Size(191, 23);
            this.sectionLoadButton.TabIndex = 0;
            this.sectionLoadButton.Text = "Load sections";
            this.sectionLoadButton.UseVisualStyleBackColor = true;
            this.sectionLoadButton.Click += new System.EventHandler(this.loadSections);
            // 
            // notebookName
            // 
            this.notebookName.Location = new System.Drawing.Point(104, 43);
            this.notebookName.Name = "notebookName";
            this.notebookName.Size = new System.Drawing.Size(100, 20);
            this.notebookName.TabIndex = 2;
            // 
            // notebookFilterLabel
            // 
            this.notebookFilterLabel.AutoSize = true;
            this.notebookFilterLabel.Location = new System.Drawing.Point(12, 46);
            this.notebookFilterLabel.Name = "notebookFilterLabel";
            this.notebookFilterLabel.Size = new System.Drawing.Size(86, 13);
            this.notebookFilterLabel.TabIndex = 3;
            this.notebookFilterLabel.Text = "Notebook name:";
            // 
            // dgv
            // 
            this.dgv.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Nodebook,
            this.Section,
            this.TimeExported,
            this.doWork});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.Location = new System.Drawing.Point(13, 69);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(1043, 429);
            this.dgv.TabIndex = 4;
            this.dgv.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_cellClick);
            // 
            // Nodebook
            // 
            this.Nodebook.Name = "Nodebook";
            this.Nodebook.Width = 300;
            // 
            // Section
            // 
            this.Section.Name = "Section";
            this.Section.Width = 300;
            // 
            // TimeExported
            // 
            this.TimeExported.Name = "Time Exported";
            this.TimeExported.Width = 300;
            // 
            // doWork
            // 
            this.doWork.FalseValue = "0";
            this.doWork.HeaderText = "Export";
            this.doWork.Name = "doWork";
            this.doWork.TrueValue = "1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(588, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(186, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Export sections to docx";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.export_clicked);
            // 
            // exportingProgressbar
            // 
            this.exportingProgressbar.Location = new System.Drawing.Point(13, 522);
            this.exportingProgressbar.Name = "exportingProgressbar";
            this.exportingProgressbar.Size = new System.Drawing.Size(1043, 23);
            this.exportingProgressbar.TabIndex = 8;
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(12, 548);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(38, 13);
            this.progressLabel.TabIndex = 9;
            this.progressLabel.Text = "Ready";
            // 
            // Mainframe
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1073, 641);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.exportingProgressbar);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.notebookFilterLabel);
            this.Controls.Add(this.notebookName);
            this.Controls.Add(this.sectionLoadButton);
            this.Name = "Mainframe";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button sectionLoadButton;
        private System.Windows.Forms.MaskedTextBox notebookName;
        private System.Windows.Forms.Label notebookFilterLabel;
        private DataGridView dgv;
        private DataGridViewTextBoxColumn Nodebook;
        private DataGridViewTextBoxColumn TimeExported;
        private DataGridViewTextBoxColumn Section;
        private DataGridViewCheckBoxColumn doWork;
        private Button button1;
        private ProgressBar exportingProgressbar;
        private Label progressLabel;
    }
}

