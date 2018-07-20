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
    class Program
    {
        static void Main(string[] args)
        {
            using (var mainDoc = WordprocessingDocument.Open(@"C:\Users\waqas.dilawar\Documents\Recruitment_Pack_–_Recruit_Specification.docx", false))
            using (var resultDoc = WordprocessingDocument.Create(@"C:\Users\waqas.dilawar\Documents\Recruitment_Pack_–_Recruit_Specification-2.docx",
              WordprocessingDocumentType.Document))
            {
                // copy parts from source document to new document
                foreach (var part in mainDoc.Parts)
                    resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                var mainDocPart = resultDoc.MainDocumentPart;
                var pages = resultDoc.ExtendedFilePropertiesPart.Properties;

                //FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                //FieldCode objFieldCode_2 = new FieldCode();
                ////it is where we need to set our pages count
                //objFieldCode_2.Text = "NUMPAGES";
                //FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
                var count = 0;
                foreach (var headerPart in resultDoc.MainDocumentPart.HeaderParts)
                {

                    
                    //Gets the text in headers
                    foreach (Text currentText in headerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {   
                     
               
                        
                        //var parentRun1 = currentText.Parent;
                        //var grandPaa1 = parentRun1.Parent;
                        //var typeOfGParent = grandPaa1.GetType();
                        //parentRun.Append(objFieldCode_2);
                        //grandPaa.Append(objRun_8);
                        //grandPaa.Append(objRun_9);
                        //grandPaa.Append(objRun_12);
                        //var typeOfParent = parentRun1.GetType();
                        // use the StringComparison parameter on methods that have it to specify how to match strings.
                        bool isReview = currentText.Text.EndsWith("Reviewed:", System.StringComparison.CurrentCultureIgnoreCase);

                        bool ignoreCaseSearchResult = currentText.Text.StartsWith(" of ", System.StringComparison.CurrentCultureIgnoreCase);
                        bool isEnds = currentText.Text.EndsWith("Page: ", System.StringComparison.CurrentCultureIgnoreCase);
                        bool isStart = currentText.Text.StartsWith("Page: ", System.StringComparison.CurrentCultureIgnoreCase);
                       
                        if (ignoreCaseSearchResult && !isEnds || isReview)
                        {
                            Regex re = new Regex(@"\d+");
                            Match m = re.Match(currentText.Text);
                            if (m.Success && !currentText.Text.Contains("Printed"))
                            {
                                var isNumeric = int.TryParse(m.Value, out int n);
                               
                                if (isNumeric)
                                {
                                    TotalCountChange.ChangeTotalCount(currentText);
                                    #region Before Changes

                                    //var parentRun = currentText.Parent;
                                    //var grandPaa = parentRun.Parent;
                                    //List<Run> runsToRemove = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(0).Take(7).ToList();
                                    //foreach (Run runToRemove in runsToRemove)
                                    //{
                                    //    if(runToRemove.InnerText!="PAGE")
                                    //    runToRemove.RemoveAllChildren();
                                    //    runToRemove.Remove();
                                    //}
                                    ////For Referencing Where to Insert Pagination Nodes
                                    //Run referenceRun = grandPaa.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Skip(0).Take(1).First();
                                    ////New Run For Pagination Node at One Line
                                    //Run newRun = new Run();


                                    #region Working Partially Fine
                                    ////RunProperties runProperties25 = new RunProperties();
                                    ////FontSize fontSize25 = new FontSize() { Val = "16" };
                                    ////runProperties25.Append(fontSize25);

                                    ////Text objText_3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                                    ////objText_3.Text = " of ";
                                    ////newRun.Append(runProperties25);
                                    ////newRun.Append(objText_3);
                                    //////grandPaa.Append(objRun_7);


                                    ////FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                                    ////newRun.Append(objFieldChar_4);
                                    //////grandPaa.Append(objRun_8);


                                    ////RunProperties runProperties5 = new RunProperties();
                                    ////Spacing spacing1 = new Spacing() { Val = -7 };
                                    ////FontSize fontSize3 = new FontSize() { Val = "16" };
                                    ////runProperties5.Append(spacing1);
                                    ////runProperties5.Append(fontSize3);
                                    ////FieldCode text3 = new FieldCode();
                                    //////it is where we need to set our pages count
                                    ////text3.Text = "NUMPAGES";

                                    ////newRun.Append(runProperties5);
                                    ////newRun.Append(text3);
                                    //////grandPaa.Append(run5);


                                    ////FieldChar objFieldChar_6 = new FieldChar() { FieldCharType = FieldCharValues.End };
                                    ////newRun.Append(objFieldChar_6);
                                    ////grandPaa.InsertBefore(newRun, referenceRun); 
                                    #endregion


                                    #region Working Perfectly Fine
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
                                    //objFieldCode_1.Text = "PAGE";
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
                                    
                                #endregion
                                }
                            }
                        }
                        //Close the Document after If and make else if out of this loop and follow below instruction
                        //Open The Document Here and Access Pages
                        #region Before Changes
                        else if (!isEnds && isStart && !currentText.Text.Contains("Printed"))
                        {
                            Regex re = new Regex(@" of \d+");
                            Match m = re.Match(currentText.Text);
                            if (m.Success)
                            {
                                re = new Regex(@"\d+");
                                m = re.Match(m.Value);
                                if (m.Success)
                                {
                                    var isNumeric = int.TryParse(m.Value, out int n);
                                    if (isNumeric)
                                    {

                                        var parentRun = currentText.Parent;
                                        var grandPaa = parentRun.Parent;



                                        #region Need to be changed
                                        //Text t = parentRun.Elements<Text>().First();
                                        //t.Text = t.Text.Substring(0, 6);
                                        ////string sub = t.Text.Substring(Math.Max(0, t.Text.Length - 5));
                                        ////t.Text= t.Text.Split(Convert.ToChar("f")).First();
                                        ////t.Text = t.Text + "f";
                                        ////t.RemoveAllChildren();
                                        ////t.Remove();
                                        ////RunProperties runProperties4 = new RunProperties();
                                        ////FontSize fontSize2 = new FontSize() { Val = "16" };

                                        ////runProperties4.Append(fontSize2);
                                        ////Text text2 = new Text();
                                        ////text2.Text = "Page:";

                                        ////parentRun.Append(runProperties4);
                                        ////parentRun.Append(text2);
                                        //#region ByWaqas

                                        //FieldChar objFieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                                        //parentRun.Append(objFieldChar_1);
                                        ////grandPaa.Append(parentRun);


                                        //RunProperties runProperties26 = new RunProperties();
                                        //FontSize fontSize26 = new FontSize() { Val = "16" };

                                        //runProperties26.Append(fontSize26);

                                        //FieldCode objFieldCode_1 = new FieldCode();
                                        //objFieldCode_1.Text = "PAGE";
                                        ////page is page number 
                                        //parentRun.Append(runProperties26);
                                        //parentRun.Append(objFieldCode_1);
                                        ////grandPaa.Append(parentRun);


                                        //FieldChar objFieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
                                        //parentRun.Append(objFieldChar_3);
                                        ////grandPaa.Append(objRun_6);

                                        //RunProperties runProperties25 = new RunProperties();
                                        //FontSize fontSize25 = new FontSize() { Val = "16" };
                                        //runProperties25.Append(fontSize25);

                                        //Text objText_3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                                        //objText_3.Text = " of ";
                                        //parentRun.Append(runProperties25);
                                        //parentRun.Append(objText_3);
                                        ////grandPaa.Append(objRun_7);


                                        //FieldChar objFieldChar_4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
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
                                        ////grandPaa.Append(objRun_12);
                                        //#endregion
                                        #endregion

                                        #region Working Perfectly Fine
                                        Text t = parentRun.Elements<Text>().First();
                                        var pageNumber = t.Text.Substring(6, 1);
                                        t.RemoveAllChildren();
                                        t.Remove();
                                        #region UnComment
                                        RunProperties runProperties4 = new RunProperties();
                                        FontSize fontSize2 = new FontSize() { Val = "16" };

                                        runProperties4.Append(fontSize2);
                                        Text text2 = new Text();
                                        text2.Text = "Page: ";

                                        parentRun.Append(runProperties4);
                                        parentRun.Append(text2);
                                        #region ByWaqas

                                        FieldChar objFieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                                        parentRun.Append(objFieldChar_1);
                                        //grandPaa.Append(parentRun);


                                        RunProperties runProperties26 = new RunProperties();
                                        FontSize fontSize26 = new FontSize() { Val = "16" };

                                        runProperties26.Append(fontSize26);

                                        FieldCode objFieldCode_1 = new FieldCode();
                                        objFieldCode_1.Text = "PAGE";
                                        //page is page number 
                                        parentRun.Append(runProperties26);
                                        parentRun.Append(objFieldCode_1);
                                        //grandPaa.Append(parentRun);


                                        FieldChar objFieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
                                        parentRun.Append(objFieldChar_3);
                                        //grandPaa.Append(objRun_6);

                                        RunProperties runProperties25 = new RunProperties();
                                        FontSize fontSize25 = new FontSize() { Val = "16" };
                                        runProperties25.Append(fontSize25);

                                        Text objText_3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                                        objText_3.Text = " of ";
                                        parentRun.Append(runProperties25);
                                        parentRun.Append(objText_3);
                                        //grandPaa.Append(objRun_7);


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
                                        //grandPaa.Append(objRun_12);
                                        #endregion
                                        #endregion

                                        #endregion
                                    }

                                }
                            }
                        }
                        #endregion

                    }
                }

            }
        }
        
        public static string GetTotalPageCount(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                var s = strSource.Substring(End);
                return s;
            }
            else
            {
                return "";
            }
        }
        public static void ApplyHeader(WordprocessingDocument doc)
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = doc.MainDocumentPart;

            HeaderPart headerPart1 = mainDocPart.AddNewPart<HeaderPart>("r97");



            Header header1 = new Header();

            Paragraph paragraph1 = new Paragraph() { };



            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Header stuff";

            run1.Append(text1);

            paragraph1.Append(run1);


            header1.Append(paragraph1);

            headerPart1.Header = header1;



            SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProperties1 == null)
            {
                sectionProperties1 = new SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "r97" };


            sectionProperties1.InsertAt(headerReference1, 0);

        }
        public static void RemoveDefaultPagination(OpenXmlElement element)
        {
            Paragraph mainParagraph = element.Elements<Paragraph>().First();
            Run mainRun = mainParagraph.Elements<Run>().ElementAt(1);
            Picture pic = mainRun.Elements<Picture>().First();
            V.Shape sh = pic.Elements<V.Shape>().First();
            V.TextBox textBox = sh.Elements<V.TextBox>().First();
            TextBoxContent textBoxContent = textBox.Elements<TextBoxContent>().First();
            Paragraph paginatedParagraph = textBoxContent.Elements<Paragraph>().First();
            //paginatedParagraph.RemoveAllChildren();
            List<Run> paginatedRuns = paginatedParagraph.Elements<Run>().ToList();
            List<Run> newRuns = paginatedRuns.GetRange(0, 7);

            var indexValue = 0;
            foreach (Run item in newRuns)
            {
                item.RemoveAllChildren();
                item.Remove();
                //RunProperties runProperties4 = new RunProperties();
                //FontSize fontSize2 = new FontSize() { Val = "16" };

                //runProperties4.Append(fontSize2);
                //Text text2 = new Text();
                //text2.Text = "P";
                //item.Append(runProperties4);
                //item.Append(text2);
                //item.Remove();
            }
        }

    }
}
