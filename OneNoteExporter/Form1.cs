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

        const string location = "C:\\Exported Sections\\";

        private String[] currentList = new String[100];

        private string XMLString="";
        private XmlDocument xmlDoc = new XmlDocument();

        private String strXML="";

        int count = 0;
        int processed = 0;


        public Mainframe()
        {
            InitializeComponent();
        }
        

        private void loadSections(object sender, EventArgs e)
        {

            GetAllSections(notebookName.Text);

            dgv.Rows.Clear();

            for(int i = 0; i < currentList.Length; i++)
            {
                String[] data = currentList[i].Split('@');
                dgv.Rows.Add(new String[] { data[0],data[1],data[3] });
            }
        }

        public void updateNotebookData()
        {
            Microsoft.Office.Interop.OneNote.Application onApplication = new Microsoft.Office.Interop.OneNote.Application();
            onApplication.GetHierarchy(System.String.Empty,
               Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out strXML);
            xmlDoc.LoadXml(strXML);
        }


        String[] GetAllSections(string notebook)
        {
            updateNotebookData();
            List<String> returnlist = new List<string>();
            var ns = XDocument.Parse(strXML).Root.Name.Namespace;
            foreach (var nodeBook in XDocument.Parse(strXML).Descendants(ns + "Notebook"))
            {
                if (nodeBook.Attribute("name").ToString().Contains(notebook)|| notebook.Equals("")||notebook.Equals(String.Empty)) { 
                    foreach (var nodeSection in nodeBook.Descendants(ns + "Section"))
                    {
                        if (nodeSection != null)
                        {
                            string s= nodeBook.Attribute("name").Value +"@"+ nodeSection.Attribute("name").Value+"@"+nodeSection.Attribute("ID").Value;
                            string path = location + nodeSection.Attribute("name").Value + ".docx";
                            if (File.Exists(path))
                            {
                                s+="@"+File.GetCreationTime(path);
                            }
                            else
                            {
                                s += "@" + "Not created";
                            }
                            returnlist.Add(s);
                        }
                    }
                }
            }
            currentList = returnlist.ToArray();
            return returnlist.ToArray();
        }

        private void dgv_cellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0&& e.RowIndex>=0)
            {
                dgv.Columns[e.ColumnIndex].ReadOnly = true;
                if (e.ColumnIndex > 2)
                {
                    dgv.Columns[e.ColumnIndex].ReadOnly = false;
                }
            }
        }

        
        private void publishDoc(string notebook,string section,string id,int position)
        {
            Microsoft.Office.Interop.OneNote.Application oneNoteInner;
            oneNoteInner = new Microsoft.Office.Interop.OneNote.Application();

            Console.WriteLine(notebook + "  " + section + "   " + id);           
            string path = section+".docx";
            if(Directory.Exists(location))
                Directory.CreateDirectory(location);
            if(File.Exists(location + path))
            {
                File.Delete(location + path);
            }
            try
            {
                oneNoteInner.Publish(id, location + path, PublishFormat.pfWord, "");
                dgv.Rows[position].Cells[2].Value = DateTime.Now;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }


        private void export_clicked(object sender, EventArgs e)
        {

            count = 0;
            for (int i = 0; i < currentList.Length; i++)
            {
                DataGridViewRow row = dgv.Rows[i];
                DataGridViewCheckBoxCell box = row.Cells[3] as DataGridViewCheckBoxCell;
                if (box.Value == "1")
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


        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int i = 0; i < currentList.Length; i++)
            {
                DataGridViewRow row = dgv.Rows[i];
                DataGridViewCheckBoxCell box = row.Cells[3] as DataGridViewCheckBoxCell;
                if (box.Value == "1")
                {
                    String[] copy = currentList[i].Split('@');
                    worker.ReportProgress(i);
                    publishDoc(copy[0], copy[1], copy[2], i);
                    worker.ReportProgress(-100);
                }

            }
            worker.ReportProgress(-1);

        }
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int value = e.ProgressPercentage;
            if (value != -100 &&value>=0)
            {
                if (value < currentList.Length) {
                    Console.WriteLine("Value passed: " + value);
                    String[] copy = currentList[value].Split('@');
                
                    progressLabel.Text = "Exporting section: " + copy[1];
                }
            }
            else
            {
                if (value == -1)
                {
                    progressLabel.Text = "Finished";
                }
                else
                {
                    exportingProgressbar.PerformStep();
                }
            }
        }


    }
}
