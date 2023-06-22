using ConsoleApp_FileChange;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static string filePath = @"C:\\sampleDoc\\TPRP_Notification.docx";
    public const string StampOnSanitizedFile = "Sanitized By Global Markets - EY Knowledge";
    static void Main()
    {
        //Generic1<Request> generic1 = new Generic1<Request>();
        //generic1.Method1<int>(new Request("tenant1"), 1);
        //Generic<string,int> generic = new Generic<string,int>();
        //generic.WriteData("string", 100);
        //ChildClass  childClass = new ChildClass();
        //childClass.DisplayMessage("Titan");
        //string name = "xxx";
        //clsB b =new clsB("ttt") { Name="vvv"};
        //b.SetValue(1000,ref name);
        //b.DisplayValue();

        //recordTest objRec=new recordTest("Abhi");
        //Console.WriteLine(b.Name);

        //PdfChange.AddStamp();
        //WordDocChange.AddStamp();
        //PptChange.AddText();

        //new PptChange().AddText();

        ////WordDocumentEditor editor = new WordDocumentEditor();
        ////editor.AddText(filePath);

        //WordProcessingTextBox.AddTextBoxToDocument(filePath);
        ////WordProcessingTextBox.PlaceTextAtCoordinate(para, "Text at 120.5,120.5", 120.1, 120.1);
    }
    #region Old code
    //static void ChangeHeader(string documentPath)
    //{
    //    // Replace header in target document with header of source document.
    //    using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
    //    {
    //        // Get the main document part
    //        MainDocumentPart mainDocumentPart = document.MainDocumentPart;

    //        // Check if the footer part already exists
    //        FooterPart footerPart = mainDocumentPart.FooterParts.FirstOrDefault();

    //        string footerPartId;

    //        if (footerPart == null)
    //        {
    //            // Create a new footer part
    //            footerPart = mainDocumentPart.AddNewPart<FooterPart>();

    //            // Generate a unique relationship ID for the footer part
    //            footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

    //            // Create the Footer reference and set the relationship ID
    //            var footerReference = new FooterReference() { Type = HeaderFooterValues.First, Id = footerPartId };

    //            // Access the section properties of the first section
    //            var sectionProperties1 = mainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();

    //            if (sectionProperties1 == null)
    //            {
    //                // Create new section properties if they don't exist
    //                sectionProperties1 = new SectionProperties();
    //                mainDocumentPart.Document.Body.AppendChild(sectionProperties1);
    //            }

    //            // Add the footer reference to the section properties
    //            sectionProperties1.AppendChild(footerReference);
    //        }
    //        footerPartId = mainDocumentPart.GetIdOfPart(footerPart);


    //        //// Delete the existing header and footer parts
    //        //mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
    //        //mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

    //        //// Create a new header and footer part
    //        //HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
    //        //FooterPart footerPart = mainDocumentPart.AddNewPart<FooterPart>();

    //        // Get Id of the headerPart and footer parts
    //        //string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
    //        //string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

    //        //GenerateHeaderPartContent(headerPart);



    //        GenerateFooterPartContent(footerPart);

    //        //// Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
    //        //IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

    //        //foreach (var section in sections)
    //        //{
    //        //    // Delete existing references to headers and footers
    //        //    section.RemoveAllChildren<HeaderReference>();
    //        //    section.RemoveAllChildren<FooterReference>();

    //        //    // Create the new header and footer reference node
    //        //    //section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
    //        //    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
    //        //}
    //    }
    //}

    //static void GenerateHeaderPartContent(HeaderPart part)
    //{
    //    Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
    //    header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
    //    header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    //    header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
    //    header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    //    header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
    //    header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
    //    header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
    //    header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
    //    header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
    //    header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    //    header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
    //    header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
    //    header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
    //    header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
    //    header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

    //    Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

    //    ParagraphProperties paragraphProperties1 = new ParagraphProperties();
    //    ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

    //    paragraphProperties1.Append(paragraphStyleId1);

    //    Run run1 = new Run();
    //    Text text1 = new Text();
    //    text1.Text = "Header";

    //    run1.Append(text1);

    //    paragraph1.Append(paragraphProperties1);
    //    paragraph1.Append(run1);

    //    header1.Append(paragraph1);

    //    part.Header = header1;
    //}

    //static void GenerateFooterPartContent(FooterPart part)
    //{
    //    Footer footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
    //    footer1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
    //    footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    //    footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
    //    footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    //    footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
    //    footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
    //    footer1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
    //    footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
    //    footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
    //    footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    //    footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
    //    footer1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
    //    footer1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
    //    footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
    //    footer1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

    //    Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

    //    ParagraphProperties paragraphProperties1 = new ParagraphProperties();
    //    ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

    //    paragraphProperties1.Append(paragraphStyleId1);

    //    Run run1 = new Run();
    //    Text text1 = new Text();
    //    text1.Text = "Footer";

    //    run1.Append(text1);

    //    paragraph1.Append(paragraphProperties1);
    //    paragraph1.Append(run1);

    //    footer1.Append(paragraph1);

    //    part.Footer = footer1;
    //}

    //static void AddFooter()
    //{
    //    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
    //    {
    //        // Get the main document part
    //        MainDocumentPart docPart = doc.MainDocumentPart;
    //        // Check if the footer part already exists
    //        FooterPart footerPart = docPart.FooterParts.FirstOrDefault();

    //        if (footerPart == null)
    //        {
    //            // Create a new footer part
    //            footerPart = docPart.AddNewPart<FooterPart>();

    //            // Generate a unique relationship ID for the footer part
    //            string footerPartId = docPart.GetIdOfPart(footerPart);

    //            // Create the Footer reference and set the relationship ID
    //            var footerReference = new FooterReference() { Type = HeaderFooterValues.First, Id = footerPartId };

    //            // Access the section properties of the first section
    //            var sectionProperties1 = docPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();

    //            if (sectionProperties1 == null)
    //            {
    //                // Create new section properties if they don't exist
    //                sectionProperties1 = new SectionProperties();
    //                docPart.Document.Body.AppendChild(sectionProperties1);
    //            }

    //            // Add the footer reference to the section properties
    //            sectionProperties1.AppendChild(footerReference);
    //        }


    //        // Create a Paragraph object
    //        Paragraph paragraph = new Paragraph();

    //        // Create a Run object with the text you want to add
    //        Run runStamp = new Run(new Text(StampOnSanitizedFile));

    //        // Create a RunProperties object
    //        RunProperties runProperties = new RunProperties();

    //        // Set the font size
    //        FontSize fontSize = new FontSize() { Val = "30" }; // 20 half-point font size
    //        runProperties.Append(fontSize);

    //        // Set the font color
    //        Color color = new Color() { Val = "FF0000" }; // Red color
    //        runProperties.Append(color);

    //        // Set the bold style
    //        Bold bold = new Bold();
    //        runProperties.Append(bold);

    //        // Set the RunProperties for the Run
    //        runStamp.Append(runProperties);

    //        // Append the Run object to the Paragraph
    //        paragraph.Append(runStamp);

    //        // Create a ParagraphProperties object
    //        ParagraphProperties paragraphProperties = new ParagraphProperties();

    //        // Create a Justification object to align the text
    //        Justification justification = new Justification() { Val = JustificationValues.Left };

    //        // Set the alignment to the left
    //        paragraphProperties.Append(justification);

    //        // Set the ParagraphProperties for the Paragraph
    //        paragraph.Append(paragraphProperties);

    //        // Append the Paragraph to the first footer part
    //        Footer footer1 = new Footer();

    //        footer1.AppendChild(paragraph);
    //        footerPart.Footer = footer1;

    //        // Get or create the SectionProperties of the first section
    //        SectionProperties sectionProperties = docPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
    //        if (sectionProperties == null)
    //        {
    //            sectionProperties = new SectionProperties();
    //            docPart.Document.Body.AppendChild(sectionProperties);
    //        }

    //        // Create or update the FooterReference for the first page
    //        var footerReference1 = sectionProperties.Elements<FooterReference>().FirstOrDefault(fr => fr.Type == HeaderFooterValues.First);
    //        if (footerReference1 == null)
    //        {
    //            footerReference1 = new FooterReference() { Type = HeaderFooterValues.First, Id = docPart.GetIdOfPart(footerPart) };
    //            sectionProperties.AppendChild(footerReference1);
    //        }
    //        else
    //        {
    //            footerReference1.Id = docPart.GetIdOfPart(footerPart);
    //        }

    //        sectionProperties.PrependChild<FooterReference>(new FooterReference() { Id = footerReference1.Id });

    //        //// Save the changes
    //        //doc.Save();

    //        //// Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
    //        //IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

    //        //foreach (var section in sections)
    //        //{
    //        //    // Delete existing references to headers and footers
    //        //    section.RemoveAllChildren<HeaderReference>();
    //        //    section.RemoveAllChildren<FooterReference>();

    //        //    // Create the new header and footer reference node
    //        //    //section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
    //        //    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
    //        //}
    //    }
    //} 
    #endregion

    public record recordTest(string name)
    {

    }

    public class clsB
    {
        private int _number;
        private string _name;
        public clsB(string name)
        {
            Name = name;
        }



        public void SetValue(int number, ref string name)
        {
            _number = number+ 1;
            _name = name = name+" Change";
        }

        public void DisplayValue()
        {
            Console.WriteLine($"{_number} {_name}");
        }


        public string Name { get; set; }        
    }
    internal class BaseClass
    {
        internal virtual void DisplayMessage(string msg)
        {
            Console.WriteLine($"Base class- {msg}");
        }
    }

    internal class ChildClass:BaseClass
    {
        internal override void DisplayMessage(string msg)
        {
            Console.WriteLine($"Child class- {msg}");
        }
    }

}
