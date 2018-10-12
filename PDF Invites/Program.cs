using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_Invites
{
    public class Program
    {
        public string topic { get; set; }
        public string front { get; set; }
        public string back { get; set; }
        public string orientation { get; set; }
        public string assembly { get; set; }
        public int x1 { get; set; }
        public int y1 { get; set; }
        public int x2 { get; set; }
        public int y2 { get; set; }
        public int font_size { get; set; }
        public int no_of_columns { get; set; }
        public string brush_color { get; set; }
        public string description { get; set; }
        public string cred { get; set; }
        public string date { get; set; }
        public string venue { get; set; }
        public string present { get; set; }
        public string mtgcd { get; set; }
        public string reg { get; set; }
        public string rsvp { get; set; }
        public string active { get; set; }
        public string page3 { get; set; }
        public string page4 { get; set; }
        public int textWidthLimit { get; set; } 
        public List<topics> all_topics { get; set; }        
    }

    public class topics
    {
        public string topic { get; set; }        
    }

}
