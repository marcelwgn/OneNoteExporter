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

using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
namespace OneNoteExporter
{
    public partial class Mainframe : Form
    {
        /*
         * Codes:
         * #C100 Adding to DGV using positions
         * 
         * #C200 Reading section list,information based on position
         * #C201 Writing section list,information based on position
         * 
         * #C300 Reading settings; Program settings based on position
         * #C301 Writing settings; Program settings based on position
         * #C310 Reading settings; Section information based on position
         * #C311 Writing settings; Section information based on position
         * 
         */


        //Values for datagridview
        public const int exportDateIndexDGV = 4;
        public const int fileNameIndexDGV = 3;
        public const int exportBoxIndexDGV = 0;
        public const int sectionGroupIndexDGV=1;
        public const int sectionIndexDGV=2;

        //Values for the currentlist String
        const int notebookIndex = 0;
        const int sectiongroupIndex = 1;
        const int sectionIndex = 2;
        const int IDIndex = 3;
        const int fileNameIndex = 4;
        const int checkIndex = 5;


        //The relative folder where the sections will be exported to
        string folderName = "WORK";



        //The sections, its IDs and the notebook they are in;
        private NotebookObject[] notebooks = new NotebookObject[0];
        private NotebookObject currentNotebook;
        private SectionObject[] allSections = new SectionObject[0];

        //Everything needed for loading sections and filtering 
        private string XMLString="";
        private XmlDocument xmlDoc = new XmlDocument();
        private String strXML="";

        //Split character used for filtering the strings
        public const char splitter = '|';

        //count of sections that need to be processed
        int count = 0;
        //count of processed sections
        int processed = 0;


        static Mainframe()
        {

        }


        public Mainframe()
        {
            SettingsManager.initSettings();
            InitializeComponent();
            loadSections(null,null);
            this.exportFolderBox.Text = SettingsManager.exportFolder;
        }


        /*
         * Loads the sections from the XML OneNote hierarchy and writes them to the string array
         * currentlist
         */
        private void loadSections(object sender, EventArgs e)
        {

            GetAllSections(sectionNameBox.Text);
            dgvNotebook.Rows.Clear();
            dgvNotebook.Refresh();
            foreach (NotebookObject notebook in notebooks)
            {
                dgvNotebook.Rows.Add(notebook.name);

            }
            bool loaded = false;
            for (int i= 0;i < notebooks.Length;i++)
            {
                if (notebooks[i].name.Equals(SettingsManager.notebook))
                {
                    loadNotebookToDGV(notebooks[i]);
                    dgvNotebook.CurrentCell = dgvNotebook.Rows[i].Cells[0];
                    currentNotebook = notebooks[i];
                    //dgvNotebook.MultiSelect = false;
                    loaded = true;
                }
            }
            if (!loaded)
            {
                loadNotebookToDGV(notebooks[0]);
            }
        }


        /*
         * Loads a given notebook objects sections to the data grid view for sections.
         */
        private void loadNotebookToDGV(NotebookObject notebook)
        {
            SettingsManager.notebook = notebook.name;
            dgvSection.Rows.Clear();
            dgvSection.Refresh();
            List<String> added = new List<string>();
            string error = "";
            bool showError = false;
            foreach (SectionObject section in notebook.sections)
            {
                string errorAdder = "";
                if (added.Contains(section.fileName))
                {
                    showError = true;
                    if (error == "")
                    {
                        error += section.section;
                    }
                    else
                    {
                        error += "," + section.section;
                    }
                    errorAdder = "    DUPLICATE EXPORT FILENAME";
                }
                //@ Adding to dgv based on position !
                //#C100
                dgvSection.Rows.Add(section.export,
                        section.sectionGroup,
                        section.section,
                        section.fileName,
                        section.exportTime
                        );
            }

            if (showError)
            {
                MessageBox.Show("The sections : \n" + error + " have export file names that were found more than once.\n" +
                         "There may be problems with exporting due to multiple sections with same name in the same notebook.");
            }

        }

        
        /*
         * Updates the notebook xml 
         */
        public void updateNotebookData()
        {
            Microsoft.Office.Interop.OneNote.Application onApplication = new Microsoft.Office.Interop.OneNote.Application();
            onApplication.GetHierarchy(System.String.Empty,
               Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out strXML);
            xmlDoc.LoadXml(strXML);
        }


        /*
         * Returns an array of all sections found,where every string is a single section. Structure of string:
         * notebook name,SPLITTER,section name,SPLITTER,ID,SPLITTER, exported file creation time,SPLITTER,FileName,
         */
        SectionObject[] GetAllSections(string section)
        {
            updateNotebookData();
            var ns = XDocument.Parse(strXML).Root.Name.Namespace;

            List<SectionObject> returnlist = new List<SectionObject>();
            //List<NotebookObject> notebooks = new List<NotebookObject>();

            notebooks = new NotebookObject[XDocument.Parse(strXML).Descendants(ns + "Notebook").Count()];

            int index = 0;

            XDocument sd=XDocument.Parse(strXML);
            foreach (var nodeBook in XDocument.Parse(strXML).Descendants(ns + "Notebook"))
            {

                NotebookObject currentNotebook = new NotebookObject(nodeBook.Attribute("name").Value);


                foreach (XElement element in nodeBook.Elements())
                {
                        
                    string type = element.Name.ToString().Split('}')[1];
                    String sectionGroupName = "";
                    Console.WriteLine(element.Attribute("name").Value);
                    SectionObject toAdd = new SectionObject();
                    if (type == "SectionGroup")
                    {
                        if (!element.Attribute("name").Value.ToString().Equals("OneNote_RecycleBin"))
                        {
                            sectionGroupName = element.Attribute("name").Value;
                            foreach (var nodeSection in element.Descendants(ns + "Section"))
                            {
                                if (nodeSection.Attribute("name").ToString().Contains(section) || section.Equals("") || section.Equals(String.Empty))
                                {
                                    toAdd = new SectionObject(nodeBook.Attribute("name").Value, sectionGroupName, nodeSection.Attribute("name").Value, nodeSection.Attribute("ID").Value);
                                    if (!nodeSection.Parent.Attribute("name").Value.ToString().Equals(toAdd.sectionGroup))
                                    {
                                        toAdd = new SectionObject(nodeBook.Attribute("name").Value, sectionGroupName, nodeSection.Parent.Attribute("name").Value.ToString() + "-" + toAdd.section, nodeSection.Attribute("ID").Value);
                                    }
                                    if (!returnlist.Contains(toAdd))
                                    {
                                        currentNotebook.addSectionTry(toAdd);
                                        returnlist.Add(toAdd);
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        if (element.Attribute("name").ToString().Contains(section) || section.Equals("") || section.Equals(String.Empty))
                        {
                            toAdd = new SectionObject(nodeBook.Attribute("name").Value, sectionGroupName, element.Attribute("name").Value, element.Attribute("ID").Value);
                            if (!returnlist.Contains(toAdd))
                            {
                                currentNotebook.addSectionTry(toAdd);
                                returnlist.Add(toAdd);
                            }
                        }
                    }
                }
                notebooks[index]=currentNotebook;
                index++;
            }
            currentNotebook = notebooks[0];
            allSections = returnlist.ToArray();
            SettingsManager.checkForUnusedSetting();
            return returnlist.ToArray();
        }


        /*
         * Checks if the clicked cell can be edited or should be left readonly
         */
        private void dgv_cellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0&& e.RowIndex>=0)
            {
                dgvSection.Columns[e.ColumnIndex].ReadOnly = true;
                if (e.ColumnIndex.ToString()==exportBoxIndexDGV.ToString()
                    ||e.ColumnIndex.ToString()==fileNameIndexDGV.ToString())
                {
                    dgvSection.Columns[e.ColumnIndex].ReadOnly = false;
                }
            }
        }


        /*
         * Buttonhandler for the export button
         */
        private void export_clicked(object sender, EventArgs e)
        {

            count = 0;
            foreach(SectionObject section in currentNotebook.sections)
            {
                if (section.export)
                {
                    count++;
                }
            }
            exportingProgressbar.Maximum = count;
            exportingProgressbar.Step = 1;
            exportingProgressbar.Value = 0;
            processed = 0;


            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerAsync();
        }


        /*
         * Exports all of the sections that have been ticked. This will be done using a backgroundworker.
         * This is the backgroundworkers doWork method.
         */
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 0; i <currentNotebook.sections.Count; i++)
            {
                if (currentNotebook.sections[i].export)
                {
                    worker.ReportProgress(i);
                    currentNotebook.sections[i].publishToDOCX();
                    worker.ReportProgress(-100);
                }

            }
            worker.ReportProgress(-1);
        }


        /*
         * Updates the progress based on the given number.
         * Any value larger -1 is the index of the section in currentlist.
         * Value -100: progressBar.PerformStep();
         * Value -1: exporting finished
         */
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int value = e.ProgressPercentage;
            if (value != -100 &&value>=0)
            {
                if (value < allSections.Length) {                
                    progressLabel.Text = "Exporting section: " 
                        + currentNotebook.sections[value].section
                        +" in section group: "
                        + currentNotebook.sections[value].sectionGroup;
                }
            }
            else
            {
                if (value == -1)
                {
                    progressLabel.Text = "Finished";
                    this.loadNotebookToDGV(currentNotebook);
                }
                else
                {
                    exportingProgressbar.PerformStep();
                }
            }
        }


        /*
         * Checks all sections for export
         */
        private void checkAllSections_Click(object sender, EventArgs e)
        {
            if (allSections.Length != 0)
            {

                for (int i = 0; i < dgvSection.RowCount; i++)
                {
                    DataGridViewRow row = dgvSection.Rows[i];
                    DataGridViewCheckBoxCell box = row.Cells[exportBoxIndexDGV] as DataGridViewCheckBoxCell;
                    box.Value = "1";
                }
            }
        }


        /*
         * Unchecks all sections for export
         */
        private void uncheckAllButton_Click(object sender, EventArgs e)
        {
            if (allSections.Length != 0) { 
            for (int i = 0; i < dgvSection.RowCount; i++)
            {
                DataGridViewRow row = dgvSection.Rows[i];
                DataGridViewCheckBoxCell box = row.Cells[exportBoxIndexDGV] as DataGridViewCheckBoxCell;
                box.Value = "0";
            }
        }
        }

        /*
         * Saves the sections and their check for export status when form is closing AND 
         * the "Save exports when closing" checkbox is checked
         */
        private void Mainframe_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (exportsClosingCheck.Checked) {
                saveSettings(null,null);
            }
        }


        /*
         * Applies the presets to data grid view
         */
        private void applyPresets_Click(object sender, EventArgs e)
        {
            foreach(SectionObject section in allSections)
            {
                SettingsManager.applySetting(section);
            }
        }


        /*
         * Saves the current exports as presets
         */
        private void saveSettings(object sender, EventArgs e)
        {
           foreach(NotebookObject notebook in notebooks)
            {
                notebook.commitToSettingsmanager();
            }
            SettingsManager.saveSettings();
        }


        /*
         * Loads the notebook that is represented by the clicked cell to the section data grid view
         */
        private void dgvNotebook_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                loadNotebookToDGV(notebooks[e.RowIndex]);
                currentNotebook = notebooks[e.RowIndex];
            }
        }


        /*
         * Finds the sectionobject whichs value has been changed and updates it with the new values
         */
        private void dgvSection_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && 
                (e.ColumnIndex.ToString() == exportBoxIndexDGV.ToString()
                    || e.ColumnIndex.ToString() == fileNameIndexDGV.ToString()) )
            {
                String section = dgvSection.Rows[e.RowIndex].Cells[sectionIndex].Value.ToString();
                String sectiongroup = dgvSection.Rows[e.RowIndex].Cells[sectionGroupIndexDGV].Value.ToString();
                String timeExported = dgvSection.Rows[e.RowIndex].Cells[exportDateIndexDGV].Value.ToString();
                foreach (SectionObject sectionObject in currentNotebook.sections)
                {
                    if (sectionObject.isThisObject(section, sectiongroup, timeExported))
                    {
                        sectionObject.updateStatus(dgvSection.Rows[e.RowIndex].Cells[fileNameIndexDGV].Value.ToString(),
                            dgvSection.Rows[e.RowIndex].Cells[exportBoxIndexDGV].Value.ToString());
                    }
                }
                
            }
        }


        /*
         * Updates the exportfolder setting in SettingManager when the user changed the 
         * text of the export masked text box
         */
        private void exportFolderBox_TextChanged(object sender, EventArgs e)
        {
            SettingsManager.exportFolder = exportFolderBox.Text;
        }
    }

}
