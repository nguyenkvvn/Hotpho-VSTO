using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Threading;

namespace Hotpho
{

    //The PGuard class handles the actual range conversion and cleansing for the plugin.
    class PGuard
    {
        
        //This method accepts a range to protect, and then returns a "protected" range
        //Kept for legacy/reference purposes.
        public static Range LEGACYprotectRange(Range rng)
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

                //insert the token
                string rr = "❤";

                //insert the token
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

        //This updated method accepts a range to protect, and then returns a "protected" range
        public static Range protectRange(Range rng)
        {
            //tRNG is the temporary range that will be modified from the given Range rng parameter
            Range tRng = rng;

            //thread it to insert tokens
            tRng.Text = insertToken("❤", tRng.Text);
            //replace tokens
            tRng = replaceToken("❤", tRng);

            return tRng;
        }

        //thread-able method to insert tokens
        public static String insertToken(string token, string s)
        {
            String tString = s;
            int sLength = s.Length;

            //swithing to Parallel.For type loop to take advantage of threading
            //consequentially, a 100% originality content score is no longer possible, as this does not guarantee that *all* characters will get token'd. (meaning there will be some "unoriginality", debunking the 100% original content red flag)
            //so the issue with the method below is that tokens are inserted inconcistiently. 
            Parallel.For(0, sLength - 1, i =>
            {
                //insert the token
                //Debug.Write("\n[NOTE] Inserting token " + token + " at index " + i + " out of " + sLength + " with character: " + tString.Substring(i, 1));
                if ((tString[i] != token[0]) && (tString[i] != ' ')) /*(!(tString.Substring(i, 1).Equals(token)) || !(tString.Substring(i,1).Equals(" ")))*/
                {
                    tString = tString.Insert(i, token);
                }
            });

            //sequential cleanup
            //the cleanup removes duplicate tokens and inserts a token that may be missed by the Parallel.For
            //it does not behave as expected in that there is a token after every single character, but it there is a token after enough.
            string dString = "";
            //working backwards
            for (int i = tString.Length - 1; i > 0; i--)
            {
                //if the character behind the current index is NOT a token
                if (tString[i - 1] != token[0])
                {
                    //and the current character is also NOT a token
                    if (tString[i] != token[0])
                    {
                        //then insert the token
                        //tString.Insert(i, token);

                        //ALT
                        dString = dString + token + tString[i];
                    }
                    //and the current character is a token
                    else
                    {
                        //then do nothing

                        //ALT
                        dString = dString + tString[i];
                    }
                }
                //if the character behind the current index IS a token
                else if (tString[i - 1] == token[0])
                {
                    //tString.Remove(i, 1);

                    //aka dont add it
                    //dString = dString + tString[i];

                    //and the current character is also a token
                    if (tString[i] == token[0])
                    {
                        //then do not insert the token

                        //ALT
                        dString = dString;
                    }
                    //and the current character is not a token
                    else
                    {
                        //then do nothing

                        //ALT
                        dString = dString + tString[i];
                    }
                }
            }

            //return the reversed reverse string
            return new string(dString.ToCharArray().Reverse().ToArray());
        }

        //method to convert tokens to random string
        //migrated from the larger legacy method
        public static Range replaceToken(String token, Range r)
        {
            Range rTmp = r;

            //borrowed from the StackOgreflow
            //the AllowedChars string is the list of characters allowed to exist in a "randomly" generated string to be inserted.
            //the Random object generates random strings based on the allowed characters.
            //the ^ string has been removed- reserved character
            const string AllowedChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz#@$*()";
            Random rngG = new Random();
            List<string> rStrings = RandomStrings(AllowedChars, 1, 16, rTmp.Text.Length, rngG).ToList();

            //set the search and replacement formatting and replacement parameters
            //filter to search all formatting
            rTmp.Find.ClearFormatting();
            //find the target token
            rTmp.Find.Text = token;
            //no special formatting for inserted text
            rTmp.Find.Replacement.ClearFormatting();
            //insert a string from the random strings list generated
            //TODO: insert *random* string from the list of strings; this code currently only selects and inserts the first one
            rTmp.Find.Replacement.Text = rStrings.First();
            //set the hidden property of the formatting
            rTmp.Find.Replacement.Font.Hidden = 1;

            //replace the token as defined in the parameters above
            object replaceAll = WdReplace.wdReplaceAll;
            //NOTE: For future reference, Type.Missing MUST BE USED with the Find.Execute() method. Defining a local missing type will not work.
            rTmp.Find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            return rTmp;
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
