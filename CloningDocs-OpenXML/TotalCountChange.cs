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
            var pageNumber = "";

            List<Run> runsToRemove = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Take(7).ToList();
            foreach (Run runToRemove in runsToRemove)
            {
                var index = runsToRemove.IndexOf(runToRemove);
                if (index == 4)
                {
                    pageNumber = runToRemove.InnerText;
                }
                runToRemove.RemoveAllChildren();
                runToRemove.Remove();

                //if (runToRemove.InnerText != "")
                //{
                //    Text textToRemove = runToRemove.Descendants<Text>().Where(d => d.LocalName == "t").FirstOrDefault();
                //    if (textToRemove == null)
                //    {
                //        var instrText = runToRemove.Descendants().Where(d => d.LocalName=="instrText").FirstOrDefault();
                //        if (instrText != null)
                //        {
                //            instrText.RemoveAllChildren();
                //            instrText.Remove();
                //        }
                //    }
                //    if (textToRemove != null)
                //    {
                //        textToRemove.RemoveAllChildren();
                //        textToRemove.Remove();
                //    }
                //}

            }
#warning Uncomment thes are changes made involving Grandpaa

            #region UnComment
            //For Referencing Where to Insert Pagination Nodes
            Run referenceRun = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(1).Take(1).First();
            //New Run For Pagination Node at One Line
            Run newRun = new Run();
            RunProperties runProperties4 = new RunProperties();
            FontSize fontSize2 = new FontSize() { Val = "16" };

            runProperties4.Append(fontSize2);
            Text text2 = new Text();
            text2.Text = "Page:";

            newRun.Append(runProperties4);
            newRun.Append(text2);
            #region ByWaqas

            FieldChar objFieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            newRun.Append(objFieldChar_1);
            //grandPaa.Append(parentRun);


            RunProperties runProperties26 = new RunProperties();
            FontSize fontSize26 = new FontSize() { Val = "16" };

            runProperties26.Append(fontSize26);

            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE ";

            //page is page number 
            newRun.Append(runProperties26);
            newRun.Append(fieldCode1);
            //grandPaa.Append(parentRun);
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };
            newRun.Append(fieldChar2);


            FieldChar objFieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
            newRun.Append(objFieldChar_3);
            //grandPaa.Append(objRun_6);

            RunProperties runProperties25 = new RunProperties();
            FontSize fontSize25 = new FontSize() { Val = "16" };
            runProperties25.Append(fontSize25);

            Text objText_3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            objText_3.Text = " of ";
            newRun.Append(runProperties25);
            newRun.Append(objText_3);
            //grandPaa.Append(objRun_7);


            FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            newRun.Append(objFieldChar_4);
            //grandPaa.Append(objRun_8);


            RunProperties runProperties5 = new RunProperties();
            Spacing spacing1 = new Spacing() { Val = -7 };
            FontSize fontSize3 = new FontSize() { Val = "16" };
            runProperties5.Append(spacing1);
            runProperties5.Append(fontSize3);

            FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode2.Text = " NUMPAGES  ";

            newRun.Append(runProperties5);
            newRun.Append(fieldCode2);
            //grandPaa.Append(run5);
            FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };
            newRun.Append(fieldChar5);

            FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
            newRun.Append(objFieldChar_6);
            //grandPaa.Append(objRun_12);
            #endregion
            grandPaa.InsertBefore(newRun, referenceRun);
            #endregion
#warning Uncomment this it is before changes made involving Grandpaa
            //foreach (Run runToRemove in runsToRemove)
            //{
            //    var index = runsToRemove.IndexOf(runToRemove);
            //    switch (index)
            //    {
            //        case 0:
            //            RunProperties runProperties1 = new RunProperties();
            //            FontSize fontSize1 = new FontSize() { Val = "16" };

            //            runProperties1.Append(fontSize1);
            //            Text text1 = new Text();
            //            text1.Text = "Page: ";

            //            parentRun.Append(runProperties1);
            //            parentRun.Append(text1);
            //            break;
            //        case 1:
            //            RunProperties runProperties2 = new RunProperties();
            //            FontSize fontSize2 = new FontSize() { Val = "16" };

            //            runProperties2.Append(fontSize2);
            //            Text text2 = new Text();
            //            text2.Text = " ";

            //            parentRun.Append(runProperties2);
            //            parentRun.Append(text2);
            //            break;
            //        case 2:
            //            FieldChar objFieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            //            parentRun.Append(objFieldChar_1);



            //            RunProperties runProperties26 = new RunProperties();
            //            FontSize fontSize26 = new FontSize() { Val = "16" };

            //            runProperties26.Append(fontSize26);

            //            FieldCode objFieldCode_1 = new FieldCode();
            //            objFieldCode_1.Text = "PAGE";
            //            //page is page number 
            //            parentRun.Append(runProperties26);
            //            parentRun.Append(objFieldCode_1);

            //            FieldChar objFieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
            //            parentRun.Append(objFieldChar_3);
            //            break;
            //        case 3:
            //            RunProperties runProperties3 = new RunProperties();
            //            FontSize fontSize03 = new FontSize() { Val = "16" };

            //            runProperties3.Append(fontSize03);
            //            Text text03 = new Text();
            //            text03.Text = " ";

            //            parentRun.Append(runProperties3);
            //            parentRun.Append(text03);
            //            break;
            //        case 4:
            //            //RunProperties runProperties4 = new RunProperties();
            //            //FontSize fontSize4 = new FontSize() { Val = "16" };

            //            //runProperties4.Append(fontSize4);
            //            //Text text4 = new Text();
            //            //text4.Text = " "+ Convert.ToInt32(pageNumber);

            //            //parentRun.Append(runProperties4);
            //            //parentRun.Append(text4);
            //            break;
            //        case 5:
            //            RunProperties runProperties6 = new RunProperties();
            //            FontSize fontSize6 = new FontSize() { Val = "16" };

            //            runProperties6.Append(fontSize6);
            //            Text text6 = new Text();
            //            text6.Text = "of ";

            //            parentRun.Append(runProperties6);
            //            parentRun.Append(text6);
            //            break;
            //        case 6:
            //            FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            //            parentRun.Append(objFieldChar_4);
            //            //grandPaa.Append(objRun_8);


            //            RunProperties runProperties5 = new RunProperties();
            //            Spacing spacing1 = new Spacing() { Val = -7 };
            //            FontSize fontSize3 = new FontSize() { Val = "16" };
            //            runProperties5.Append(spacing1);
            //            runProperties5.Append(fontSize3);
            //            FieldCode text3 = new FieldCode();
            //            //it is where we need to set our pages count
            //            text3.Text = "NUMPAGES";

            //            parentRun.Append(runProperties5);
            //            parentRun.Append(text3);
            //            //grandPaa.Append(run5);


            //            FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
            //            parentRun.Append(objFieldChar_6);
            //            break;

            //        default:
            //            break;
            //    }
            //}
            //currentText.Text = currentText.Text.Substring(0, 4);
            //    FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            //parentRun.Append(objFieldChar_4);
            ////grandPaa.Append(objRun_8);


            //RunProperties runProperties5 = new RunProperties();
            //Spacing spacing1 = new Spacing() { Val = -7 };
            //FontSize fontSize3 = new FontSize() { Val = "16" };
            //runProperties5.Append(spacing1);
            //runProperties5.Append(fontSize3);
            //FieldCode text3 = new FieldCode();
            ////it is where we need to set our pages count
            //text3.Text = "NUMPAGES";

            //parentRun.Append(runProperties5);
            //parentRun.Append(text3);
            ////grandPaa.Append(run5);


            //FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
            //parentRun.Append(objFieldChar_6);
        }
    }
}
