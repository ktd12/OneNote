using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;

namespace CreateOneNotePage
{

    public class clsOneNoteSection : IEquatable<clsOneNoteSection>
    {

        public string name { get; set; }

        public string id { get; set; }
  
        /// <summary>
        /// endpoint for pages
        /// </summary>
        public string pagesUrl { get; set; }

        public bool Equals(clsOneNoteSection other)
        {
            return this.name.ToUpper() == other.name.ToUpper();
        }
    }
}
