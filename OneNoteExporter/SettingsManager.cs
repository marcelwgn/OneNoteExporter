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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneNoteExporter
{
    /*
     * Holds all settings and information that need to be saved
     */
    static class SettingsManager
    {

        private static bool settingsFile = false;

        static List<SettingsHolder> settings = new List<SettingsHolder>();

        public static String notebook = "";
        public static String section = "";
        public static String exportFolder = "WORK";


        //the settings for the list (which were exported last time)
        const string settingsLocation = "OneNoteExporter.ini";


        static string executePath;


        //The absolute folder where the sections will be exported to
        public static string location
        {
            get
            {
                return executePath + "\\" + SettingsManager.exportFolder + "\\";
            }
        }

        const char splitter = '|';

        const int settingsNotebook = 0;
        const int settingsSection = 2;
        const int settingsSectionGroup = 1;
        const int settingsFileNameIndex = 3;
        const int settingsCheckIndex = 4;




        static SettingsManager()
        {
            initSettings();
            String[] args = System.Environment.GetCommandLineArgs();
            executePath = System.IO.Path.GetDirectoryName(args[0]);
        }

        /*
         * Class for representing a setting for a section
         */
        class SettingsHolder
        {
            public string section;
            public string sectionGroup;
            public string notebook;
            public string filename;
            public bool export;

            public bool used = false;

            public SettingsHolder(string _notebook, string _sectiongroup, string _section, string _filename, string export)
            {
                this.sectionGroup = _sectiongroup;
                this.section = _section;
                this.notebook = _notebook;
                this.filename = _filename;
                if (export == "1")
                {
                    this.export = true;
                }
                else
                {
                    this.export = false;
                }
            }

            public SettingsHolder(SectionObject section)
            {
                this.sectionGroup = section.sectionGroup;
                this.section = section.section;
                this.notebook = section.notebook;
                this.filename = section.fileName;
                
                    this.export = section.export;
               
            }

            public SettingsHolder(String s)
            {
                String[] data = s.Split(splitter);
                this.section = data[settingsSection];
                this.sectionGroup = data[settingsSectionGroup];
                this.notebook = data[settingsNotebook];
                this.filename = data[settingsFileNameIndex];
                if (data[settingsCheckIndex] == "1")
                {
                    this.export = true;
                }
                else
                {
                    this.export = false;
                }
            }

            public void adaptToItem(SectionObject item)
            {
                if (this.notebook.Equals(item.notebook) && this.sectionGroup.Equals(item.sectionGroup) && this.section.Equals(item.section))
                {
                    this.filename = item.fileName;
                    this.export = item.export;
                }
            }

        }

        /*
         * Tries to apply a setting otherwise applies settings based on a default structure
         */
        public static void applySetting(SectionObject section)
        {
            bool hasSetting = false;
            foreach(SettingsHolder setting in settings)
            {
                if (setting.notebook.Equals(section.notebook) &&
                    setting.sectionGroup.Equals(section.sectionGroup) &&
                    setting.section.Equals(section.section))
                {
                    section.fileName = setting.filename;
                    section.export = setting.export;
                    hasSetting = true;
                    setting.used = true;
                }
            }
            if (!hasSetting)
            {
                if (section.sectionGroup != "")
                {
                    section.fileName = section.sectionGroup + "-" + section.section;
                }
                else
                {
                    section.fileName = section.section;
                }
                section.export = false;
                SettingsHolder settingsHolder = new SettingsHolder(section.notebook,section.sectionGroup, section.section, section.fileName, "0");
                settingsHolder.used = true;
                settings.Add(settingsHolder);
            }
        }

        /*
         * Checks for settings that not have been used and creates an error message to inform the user about that
         */
        public static void checkForUnusedSetting()
        {
            String errorList = "";
            String errorListExported = "";
            foreach(SettingsHolder holder in settings)
            {
                if (!holder.used)
                {
                    errorList +="   " + holder.section + "\n";
                    if (holder.export)
                    {
                        errorListExported += "   " + holder.section + "\n";
                    }
                }
            }
            if (errorList != "")
            {
                String message = "The following sections have not been found through OneNote, but have been found in the settings: \n" + errorList;
                if (errorListExported != "")
                {
                    message += "\nThe following sections are marked as export but could not be found through OneNote: \n" + errorListExported;
                }
                MessageBox.Show(message);
            }
        }


        /*
         * Loads the the settings
         */
        public static void initSettings()
        {
            if (File.Exists(settingsLocation))
            {
                settingsFile = true;
                int counter = 1;
                string line;
                var lineCount = File.ReadLines(settingsLocation).Count();
                settings = new List<SettingsHolder>(lineCount);
                System.IO.StreamReader file =
                     new System.IO.StreamReader(settingsLocation);
                String settingsProgram = file.ReadLine();
                String[] start = settingsProgram.Split(splitter);
                notebook = start[0];
                section = start[1];
                exportFolder = start[2];
                while ((line = file.ReadLine()) != null)
                {
                    if (line != String.Empty)
                    {
                        settings.Add(new SettingsHolder(line));
                        counter++;
                    }
                }
                file.Close();
            }
        }


        public static void updateSetting(SectionObject section)
        {
            bool flag = false;
            foreach(SettingsHolder holder in settings)
            {
                if (holder.notebook.Equals(section.notebook)
                    && holder.sectionGroup.Equals(section.sectionGroup) 
                    && holder.section.Equals(section.section))
                {
                    holder.filename = section.fileName;
                    holder.export = section.export;
                    flag = true;
                }
            }
            if (!flag)
            {
                settings.Add(new SettingsHolder(section));
            }
        }
        public static void saveSettings()
        {
            if (!File.Exists(settingsLocation))
            {
                System.IO.Stream s = File.Create(settingsLocation);
                s.Close();
            }
            FileStream fs = new FileStream(settingsLocation, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            Boolean hasWriteAcces = fs.CanWrite;
            if (hasWriteAcces)
            {
                fs.Close();
                int count = 0;
                foreach(SettingsHolder holder in settings)
                {
                    if (holder.export)
                    {
                        count++;
                    }
                }
                String[] settingsIntern = new String[count + 1];
                settingsIntern[0] = notebook + splitter + section + splitter + exportFolder;
                int processed = 0;
                for (int i = 0; i < settings.Count; i++)
                {
                    String export = "0";
                    if (settings[i].export)
                    {
                        export = "1";

                        settingsIntern[processed + 1] = settings[i].notebook + splitter.ToString() +
                            settings[i].sectionGroup + splitter.ToString() +
                            settings[i].section + splitter.ToString() +
                            settings[i].filename + splitter.ToString() +
                            export;
                        processed++;
                    }
                }
                try
                {
                    File.WriteAllLines(settingsLocation, settingsIntern);
                    MessageBox.Show("Settings saved succesfully!");
                }
                catch (IOException ex)
                {
                    MessageBox.Show(ex.Message + " Try closing processes that may access this path.");
                }
            }
        }

     }
}
