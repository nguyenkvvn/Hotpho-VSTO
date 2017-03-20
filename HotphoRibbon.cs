using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace Hotpho
{
    public partial class HotphoRibbon
    {
        private void HotphoRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Range rng = document.Application.Selection.Range;

            Debug.WriteLine(rng.Text);

            Range replaceR = PGuard.protectRange(rng);

            Debug.WriteLine(replaceR.Text);

            rng = replaceR;

            //begin formatting and clensing

            Microsoft.Office.Interop.Word.Document targetDoc = Globals.ThisAddIn.Application.ActiveDocument;

            /*string startC = "";
            string endC = "";

            Find replaceMe = Globals.ThisAddIn.Application.Selection.Find;
            replaceMe.ClearFormatting();
            replaceMe.Text = "#";

            replaceMe.Replacement.ClearFormatting();
            replaceMe.Replacement.Text = "?";
            replaceMe.Replacement.Font.Size = 1;

            object replaceAll = WdReplace.wdReplaceAll;
            object missing = null;
            replaceMe.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);*/

            //Globals.ThisAddIn.Application.Selection.Find.Execute(ref startC);
        }
    }
}
