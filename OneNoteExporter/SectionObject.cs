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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneNoteExporter
{
    /*
     * Class representing a OneNote section.
     * This class has all the fields that are necessary and available through the XML of the OneNote hierarchy.
     * 
     */
    public class SectionObject
    {

        public string notebook{ get; set; }
        public string sectionGroup { get; set; }
        public string section { get; set; }
        public string ID { get; set; }

        public string fileName;
        public string exportTime;
        public bool export;


        public SectionObject(string _notebook, string _sectiongroup, string _section, string _ID)
        {
            this.notebook = _notebook;
            this.sectionGroup = _sectiongroup;
            this.section = _section;
            this.ID = _ID;
            this.fileName = "";
            this.exportTime = "";
            this.export = false;
            SettingsManager.applySetting(this);
            this.updateTime();
        }

        public SectionObject()
        {
            this.notebook ="";
            this.sectionGroup = "";
            this.section = "";
            this.ID = "";
        }


        /*
         * Updates the export time of this sectionobject based on the filename
         */
        internal void updateTime()
        {
            Console.WriteLine(SettingsManager.location + this.fileName+".docx");
            if (File.Exists(SettingsManager.location+this.fileName + ".docx"))
            {
                this.exportTime = File.GetLastWriteTime(SettingsManager.location + this.fileName + ".docx").ToString();
            }
            else
            {
                this.exportTime = "Not created";
            }
        }


        /*
         * Updates the filename and export status of this section object
         */
        public void updateStatus(string _filename,string _export)
        {
            this.fileName = _filename;
            if (_export == "1")
            {
                this.export = true;
            }
            else
            {
                this.export = false;
            }
        }


        /*
         * Publishes this section as a Word document / .DOCX
         */
        public void publishToDOCX()
        {
            //folderName = "\\" + exportFolderBox.Text + "\\";
            Microsoft.Office.Interop.OneNote.Application oneNoteInner;
            oneNoteInner = new Microsoft.Office.Interop.OneNote.Application();

            string path = this.fileName + ".docx";
            if (Directory.Exists(SettingsManager.location))
                Directory.CreateDirectory(SettingsManager.location);
            if (File.Exists(SettingsManager.location + path))
            {
                File.Delete(SettingsManager.location + path);
            }
            try
            {
                Console.WriteLine(SettingsManager.location + path);
                oneNoteInner.Publish(this.ID, SettingsManager.location + path, PublishFormat.pfWord, "");
                this.exportTime = DateTime.Now.ToString();
            }
            catch (Exception ex)
            {
                this.exportTime = "Failed with exception code: " + ex.HResult;

                MessageBox.Show(ex.Message + "  \n  while trying to create file: " + SettingsManager.location + path);
            }
        }


        /*
         * Returns true when the given section,sectiongroup and export time represent this section
         */
        public Boolean isThisObject(String section,String sectiongroup,String timeExported)
        {
            return this.section.Equals(section) && this.sectionGroup.Equals(sectiongroup) && this.exportTime.Equals(timeExported);
        }


        /*
         * Lets the SettingsManager update the settings of this sectionobject
         */
        public void commitToSettingsmanager()
        {
            SettingsManager.updateSetting(this);
        }
    }
}
