using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_Invites
{
    public static class Parameters
    {
        public static bool  _staticRun = true;
        public static string _jobNo = "ACJAN401";
        public static string _NFPId = "2018-02337";

        public static string speakerOrigText = //"Asst Professor, Manager of Clinical Research";
        "Associate Professor of Clinical Medicine, Associate Attending Physician\n" +
        "Cornell Weill Medical College, NY Presbyterian Hospital and Hospital for Special Surgery in NY";
        //"Dermatology Office of Dr. Joseph Schwartz";

        public static string speakerNewText = //"Asst Professor, \nManager of Clinical Research";
        "Associate Professor of Clinical Medicine, \nAssociate Attending Physician\n" +
        "Cornell Weill Medical College, \nNY Presbyterian Hospital and Hospital\n for Special Surgery in NY";
        //"Dermatology Office of \nDr. Joseph Schwartz";

        public static string venueOrigText = "\n123 East Cermak Road, Suite 300 Commodore & Hudson Ballrooms(.3 Miles from the Convention Center)\n";
        public static string venueNewText =  "\n123 East Cermak Road, Suite 300 \nCommodore & Hudson Ballrooms\n(.3 Miles from the Convention Center)\n";


        public static bool staticRun
        {
            get { return _staticRun; }
        }

        public static string JobNo
        {
            get { return _jobNo; }            
        }

        public static string NFPId
        {
            get { return _NFPId; }
        }

        public static string speakerText(string speakerBlock)
        {
            return speakerBlock.Replace(speakerOrigText,speakerNewText);
        }

        public static string venueText(string venueBlock)
        {
            return venueBlock.Replace(venueOrigText, venueNewText);
            //return "\nHilton Garden Inn Chicago McCormick Place\n123 East Cermak Road, Suite 300 \nCommodore & Hudson Ballrooms\n(.3 Miles from the Convention Center)\nChicago, IL 60616"; //venueBlock.Replace(venueOrigText.Trim(), venueNewText);
        }
    }
}
