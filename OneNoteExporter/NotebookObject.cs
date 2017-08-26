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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteExporter
{
    /*
     * Represents a OneNote notebook. 
     * An object of this class need to be filled with SectionObjects. This can be done 
     * using the addSectionTry method which checks if the section is in this notebook
     */
    class NotebookObject
    {
        //Name of this notebook
        public string name { get; set; }

        //List of sections held by this notebook
        public List<SectionObject> sections = new List<SectionObject>();

        public NotebookObject(string _name)
        {
            this.name = _name;
        }

        /*
         * Adds the section object to this notebook object if the sections notebook 
         * has the same name as this notebook object
         */
        public bool addSectionTry(SectionObject section)
        {
            if (section.notebook.Equals(this.name))
            { 
                sections.Add(section);
                return true;
            }
            else
            {
                return false;
            }
        }

        /*
         * Applies the settings of this notebooks sections to the SettingsManager
         */
        public void commitToSettingsmanager()
        {
            foreach(SectionObject s in sections)
            {
                s.commitToSettingsmanager();
            }
        }

    }
}
