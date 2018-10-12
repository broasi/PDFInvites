using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;

namespace PDF_Invites
{
    public class stringManipulation
    {
        public static string WordWrap(string phrase, int limit)
        {
            string[] words = phrase.Split(' ');
            StringBuilder newSentence = new StringBuilder();
            string line = "";
            foreach (string word in words)
            {
                if ((line + word).Length > limit)
                {
                    newSentence.AppendLine(line);
                    line = "";
                }

                line += string.Format("{0} ", word);
            }
            if (line.Length > 0)
                newSentence.AppendLine(line);

            return newSentence.ToString();
        }

        public static string WordWrap2(string phrase, int limit)
        {
            var _string = "Let me be the one to break you up so I have nothing to feel";

            string[] myStr = _string.Split(' ');

            Console.WriteLine(String.Join(" ", myStr.Take(3)) );

            return String.Join(" ", myStr.Take(3));
        }
    }
}


