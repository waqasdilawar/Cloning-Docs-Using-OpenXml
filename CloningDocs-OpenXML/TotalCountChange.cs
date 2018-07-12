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
            //For Referencing Where to Insert Pagination Nodes
            Run referenceRun = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(1).Take(1).First();


            Run run1 = new Run();
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "Page ";

            run1.Append(text1);

            Run run2 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            runProperties1.Append(bold1);
            runProperties1.Append(boldComplexScript1);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript1);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(runProperties1);
            run2.Append(fieldChar1);

            Run run3 = new Run();

            RunProperties runProperties2 = new RunProperties();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();

            runProperties2.Append(bold2);
            runProperties2.Append(boldComplexScript2);
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE ";

            run3.Append(runProperties2);
            run3.Append(fieldCode1);

            Run run4 = new Run();

            RunProperties runProperties3 = new RunProperties();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            runProperties3.Append(bold3);
            runProperties3.Append(boldComplexScript3);
            runProperties3.Append(fontSize2);
            runProperties3.Append(fontSizeComplexScript2);
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(runProperties3);
            run4.Append(fieldChar2);

            Run run5 = new Run();

            RunProperties runProperties4 = new RunProperties();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            NoProof noProof1 = new NoProof();

            runProperties4.Append(bold4);
            runProperties4.Append(boldComplexScript4);
            runProperties4.Append(noProof1);
            Text text2 = new Text();
            text2.Text = "2";

            run5.Append(runProperties4);
            run5.Append(text2);

            Run run6 = new Run();

            RunProperties runProperties5 = new RunProperties();
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            runProperties5.Append(bold5);
            runProperties5.Append(boldComplexScript5);
            runProperties5.Append(fontSize3);
            runProperties5.Append(fontSizeComplexScript3);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties5);
            run6.Append(fieldChar3);

            Run run7 = new Run();
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = " of ";

            run7.Append(text3);

            Run run8 = new Run();

            RunProperties runProperties6 = new RunProperties();
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(bold6);
            runProperties6.Append(boldComplexScript6);
            runProperties6.Append(fontSize4);
            runProperties6.Append(fontSizeComplexScript4);
            FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run8.Append(runProperties6);
            run8.Append(fieldChar4);

            Run run9 = new Run();

            RunProperties runProperties7 = new RunProperties();
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();

            runProperties7.Append(bold7);
            runProperties7.Append(boldComplexScript7);
            FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode2.Text = " NUMPAGES  ";

            run9.Append(runProperties7);
            run9.Append(fieldCode2);

            Run run10 = new Run();

            RunProperties runProperties8 = new RunProperties();
            Bold bold8 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(bold8);
            runProperties8.Append(boldComplexScript8);
            runProperties8.Append(fontSize5);
            runProperties8.Append(fontSizeComplexScript5);
            FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run10.Append(runProperties8);
            run10.Append(fieldChar5);

            Run run11 = new Run();

            RunProperties runProperties9 = new RunProperties();
            Bold bold9 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            NoProof noProof2 = new NoProof();

            runProperties9.Append(bold9);
            runProperties9.Append(boldComplexScript9);
            runProperties9.Append(noProof2);
            Text text4 = new Text();
            text4.Text = "2";

            run11.Append(runProperties9);
            run11.Append(text4);

            Run run12 = new Run();

            RunProperties runProperties10 = new RunProperties();
            Bold bold10 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(bold10);
            runProperties10.Append(boldComplexScript10);
            runProperties10.Append(fontSize6);
            runProperties10.Append(fontSizeComplexScript6);
            FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run12.Append(runProperties10);
            run12.Append(fieldChar6);
            
            grandPaa.InsertBefore(run1, referenceRun);
            grandPaa.InsertBefore(run2, referenceRun);
            grandPaa.InsertBefore(run3, referenceRun);
            grandPaa.InsertBefore(run4, referenceRun);
            //grandPaa.InsertBefore(run5, referenceRun);
            grandPaa.InsertBefore(run6, referenceRun);
            grandPaa.InsertBefore(run7, referenceRun);
            grandPaa.InsertBefore(run8, referenceRun);
            grandPaa.InsertBefore(run9, referenceRun);
            grandPaa.InsertBefore(run10, referenceRun);
            //grandPaa.InsertBefore(run11, referenceRun);
            grandPaa.InsertBefore(run12, referenceRun);
            //List<Run> runsToRemove = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Take(7).ToList();
            //foreach (Run runToRemove in runsToRemove)
            //{
            //    //var index = runsToRemove.IndexOf(runToRemove);
            //    //if (index == 4)
            //    //{
            //    //    pageNumber = runToRemove.InnerText;
            //    //}
            //    runToRemove.RemoveAllChildren();
            //    runToRemove.Remove();

            //    //if (runToRemove.InnerText != "")
            //    //{
            //    //    Text textToRemove = runToRemove.Descendants<Text>().Where(d => d.LocalName == "t").FirstOrDefault();
            //    //    if (textToRemove == null)
            //    //    {
            //    //        var instrText = runToRemove.Descendants().Where(d => d.LocalName=="instrText").FirstOrDefault();
            //    //        if (instrText != null)
            //    //        {
            //    //            instrText.RemoveAllChildren();
            //    //            instrText.Remove();
            //    //        }
            //    //    }
            //    //    if (textToRemove != null)
            //    //    {
            //    //        textToRemove.RemoveAllChildren();
            //    //        textToRemove.Remove();
            //    //    }
            //    //}

            //}
#warning Uncomment thes are changes made involving Grandpaa

            #region UnComment
            //For Referencing Where to Insert Pagination Nodes
            //Run referenceRun = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(1).Take(1).First();
            ////New Run For Pagination Node at One Line
            //Run newRun = new Run();
            //RunProperties runProperties4 = new RunProperties();
            //FontSize fontSize2 = new FontSize() { Val = "16" };

            //runProperties4.Append(fontSize2);
            //Text text2 = new Text();
            //text2.Text = "Page:";

            //newRun.Append(runProperties4);
            //newRun.Append(text2);
            //#region ByWaqas

            //FieldChar objFieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            //newRun.Append(objFieldChar_1);
            ////grandPaa.Append(parentRun);


            //RunProperties runProperties26 = new RunProperties();
            //FontSize fontSize26 = new FontSize() { Val = "16" };

            //runProperties26.Append(fontSize26);

            //FieldCode objFieldCode_1 = new FieldCode();
            //objFieldCode_1.Text = "PAGE" + " " + Convert.ToInt32(pageNumber);
            ////page is page number 
            //newRun.Append(runProperties26);
            //newRun.Append(objFieldCode_1);
            ////grandPaa.Append(parentRun);


            //FieldChar objFieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
            //newRun.Append(objFieldChar_3);
            ////grandPaa.Append(objRun_6);

            //RunProperties runProperties25 = new RunProperties();
            //FontSize fontSize25 = new FontSize() { Val = "16" };
            //runProperties25.Append(fontSize25);

            //Text objText_3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //objText_3.Text = " of ";
            //newRun.Append(runProperties25);
            //newRun.Append(objText_3);
            ////grandPaa.Append(objRun_7);


            //FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            //newRun.Append(objFieldChar_4);
            ////grandPaa.Append(objRun_8);


            //RunProperties runProperties5 = new RunProperties();
            //Spacing spacing1 = new Spacing() { Val = -7 };
            //FontSize fontSize3 = new FontSize() { Val = "16" };
            //runProperties5.Append(spacing1);
            //runProperties5.Append(fontSize3);
            //FieldCode text3 = new FieldCode();
            ////it is where we need to set our pages count
            //text3.Text = "NUMPAGES";

            //newRun.Append(runProperties5);
            //newRun.Append(text3);
            ////grandPaa.Append(run5);


            //FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
            //newRun.Append(objFieldChar_6);
            ////grandPaa.Append(objRun_12);
            //#endregion
            //grandPaa.InsertBefore(newRun, referenceRun); 
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
