using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Processing document");

            //Find File
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filepath = Path.Combine(desktopPath + @"\WordFile\test_xml.docx");

            //Save Original File
            string originalFilename = Path.GetFileNameWithoutExtension(filepath);
            var originalFile = Path.Combine(desktopPath, @"WordFile\" + originalFilename + "-v1.docx");
            System.IO.File.Copy(filepath, originalFile);

            //Load File
            WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true);
            Body body = doc.MainDocumentPart.Document.Body;

            //Find Title Paragraph
            IEnumerable<Paragraph> paragraphs = body.Elements<Paragraph>();
            string title = "A method to work with Mri devices";
            Paragraph titleParagraph = null;
            foreach (Paragraph p in paragraphs)
            {
                string text = p.InnerText;
                if (text.Equals(title)) {
                    titleParagraph = p;
                }
            }

            //Replace Title Text
            string correctedTitle = "A Method to Work with MRI Devices";
            foreach (Run r in titleParagraph.Elements<Run>())
            {
                foreach(Text t in r.Elements<Text>())
                {
                    t.Text = "";
                }
            }
            Run run = titleParagraph.AppendChild(new Run());
            run.AppendChild(new Text(correctedTitle));

            //Apply Paragraph Style
            ApplyStyleToParagraph(doc, "H1", "Heading1", titleParagraph);

            //Close and Save File
            doc.Save();
            doc.Close();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        public static void ApplyStyleToParagraph(WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
        {
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            StyleDefinitionsPart part =
                doc.MainDocumentPart.StyleDefinitionsPart;

            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
                AddNewStyle(part, styleid, stylename);
            }
            else
            {
                if (IsStyleIdInDocument(doc, styleid) != true)
                {
                    string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, styleid, stylename);
                    }
                    else
                        styleid = styleidFromName;
                }
            }

            pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };
        }

        public static bool IsStyleIdInDocument(WordprocessingDocument doc,
            string styleid)
        {
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style == null)
                return false;

            return true;
        }

        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart,
            string styleid, string stylename)
        {
            Styles styles = styleDefinitionsPart.Styles;

            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true
            };
            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { Val = "000000" };
            RunFonts font1 = new RunFonts() { Ascii = "Times New Roman" };
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "30" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification centerHeading = new Justification() { Val = JustificationValues.Center };
            OnOffValue hyphenOff = new OnOffValue(false);
            AutoHyphenation hyphenation = new AutoHyphenation() { Val =  hyphenOff };
            paragraphProperties.Append(centerHeading);
            paragraphProperties.Append(hyphenation);

            style.Append(styleRunProperties1);
            styles.Append(style);
            styles.Append(paragraphProperties);
        }

        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }
    }
}
