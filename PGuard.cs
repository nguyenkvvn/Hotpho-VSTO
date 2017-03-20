using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace Hotpho
{

    //The PGuard class handles the actual range conversion and cleansing for the plugin.
    class PGuard
    {
        
        //This function accepts a range to protect, and then returns a "protected" range
        public static Range protectRange(Range rng)
        {
            //tRNG is the temporary range that will be modified from the given Range rng parameter
            Range tRng = rng;
            int tRngTextLength = tRng.Text.Length;

            //borrowed from the StackOgreflow
            //the AllowedChars string is the list of characters allowed to exist in a "randomly" generated string to be inserted.
            //the Random object generates random strings based on the allowed characters.
            const string AllowedChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz#@$^*()";
            Random rngG = new Random();
            List<string> rStrings = RandomStrings(AllowedChars, 1, 16, tRngTextLength, rngG).ToList();


            //for each character in the range, insert a "target" character to replace.
            //the following for loop only handles inserting a random character. Formatting is handled outside of the loop.
            //NOTE: This is a basic for loop that can be modified for variable text lengths, different strings from the list, etc.
            //(contd.) Without threading, large documents will take an absurd amount of time to run.
            //the -1 is to prevent new lines from being appended on.
            for (int i = tRng.Text.Length - 1; i > 0 ; i--)
            {
                Debug.WriteLine("Character Number: " + i + " of " + tRngTextLength);

                //generate a random range
                //String rr = "random range";
                //start of text char and end of text char
                string rr = "❤"/*"" + rStrings.ElementAt(tRngTextLength - i) + ""*/;

                //insert the random range
                tRng.Text = tRng.Text.Insert(i, rr);
                //i = i + rr.Length;
            }

            //set the search and replacement formatting and replacement parameters
            //filter to search all formatting
            tRng.Find.ClearFormatting();
            //find the target token
            tRng.Find.Text = "❤";
            //no special formatting for inserted text
            tRng.Find.Replacement.ClearFormatting();
            //insert a string from the random strings list generated
            //TODO: insert *random* string from the list of strings; this code currently only selects and inserts the first one
            tRng.Find.Replacement.Text = rStrings.First();
            //set the hidden property of the formatting
            tRng.Find.Replacement.Font.Hidden = 1;

            //replace the token as defined in the parameters above
            object replaceAll = WdReplace.wdReplaceAll;
            //NOTE: For future reference, Type.Missing MUST BE USED with the Find.Execute() method. Defining a local missing type will not work.
            tRng.Find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //return the completed string
            return tRng;
        }

        //random string generator from StackOverflow
        private static IEnumerable<string> RandomStrings(
        string allowedChars,
        int minLength,
        int maxLength,
        int count,
        Random rng)
        {
            char[] chars = new char[maxLength];
            int setLength = allowedChars.Length;

            while (count-- > 0)
            {
                int length = rng.Next(minLength, maxLength + 1);

                for (int i = 0; i < length; ++i)
                {
                    chars[i] = allowedChars[rng.Next(setLength)];
                }

                yield return new string(chars, 0, length);
            }
        }
    }
}
