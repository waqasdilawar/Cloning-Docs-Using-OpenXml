using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using System;
using System.Linq;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CloningDocs_OpenXML
{
    class TotalCountChange
    {
       public static void ChangeTotalCount(Text currentText)
        {
            var parentRun = currentText.Parent;
            var grandPaa = parentRun.Parent;
            List<Run> runsToRemove = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(0).Take(7).ToList();
            foreach (Run runToRemove in runsToRemove)
            {
                if (runToRemove.InnerText == "8")
                {
                    runToRemove.RemoveAllChildren();
                    runToRemove.Remove();
                }
            }
            currentText.Text = currentText.Text.Substring(0, 4);
            FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            parentRun.Append(objFieldChar_4);
            //grandPaa.Append(objRun_8);


            RunProperties runProperties5 = new RunProperties();
            Spacing spacing1 = new Spacing() { Val = -7 };
            FontSize fontSize3 = new FontSize() { Val = "16" };
            runProperties5.Append(spacing1);
            runProperties5.Append(fontSize3);
            FieldCode text3 = new FieldCode();
            //it is where we need to set our pages count
            text3.Text = "NUMPAGES";

            parentRun.Append(runProperties5);
            parentRun.Append(text3);
            //grandPaa.Append(run5);


            FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
            parentRun.Append(objFieldChar_6);
        }
    }
}
