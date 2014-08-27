using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;

namespace CreateOneNotePage
{

    public class clsOneNoteNotebook: IEquatable<clsOneNoteNotebook>
    {

        public string name { get; set; }
 
        public string id { get; set; }

        /// <summary>
        /// endpoint for sections
        /// </summary>
        public string sectionsUrl { get; set; }
  
        /// <summary>
        /// URL to launch OneNote rich client
        /// </summary>
        public string oneNoteClientUrl { get; set; }
 
        /// <summary>
        /// URL to launch OneNote web experience
        /// </summary>
        public string oneNoteWebUrl { get; set; }


        public bool Equals(clsOneNoteNotebook other)
        {
            return this.name.ToUpper() == other.name.ToUpper();
        }
    }
}
