using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

namespace ConsoleApp_FileChange
{
    internal class WordDocChange
    {
        private const string FilePath = @"C:\sampleDoc\docx\TPRP_Notification.docx";
        private const string FileOutputPath = @"C:\sampleDoc\docx\TPRP_Notification-Output.docx";
        private const string StampText = "Sanitized By Global Markets - EY Knowledge";
        private const double ShapeStampWidth = 4 * 72;
        private const double ShapeStampHeight = 2 * 72;
        private const double ShapeStampXCoordinate = 50;
        private const double ShapeStampBottomPosition = 50;
        internal static void AddStamp()
        {
            // Load the document
            Document doc = new Document(FilePath);

            // Get the first page
            PageSetup pageSetup = doc.FirstSection.PageSetup;

            // Create a DocumentBuilder object
            DocumentBuilder builder = new DocumentBuilder(doc);            

            // Calculate the top position based on the page height
            double top = pageSetup.PageHeight - ShapeStampBottomPosition;

            // Create a text box shape
            Shape textBox = builder.InsertShape(ShapeType.TextBox,RelativeHorizontalPosition.Page, ShapeStampXCoordinate, RelativeVerticalPosition.Page,
                top, ShapeStampWidth, ShapeStampHeight, WrapType.None);
            textBox.Stroked = false;

            // Create a paragraph inside the text box
            Paragraph para = new Paragraph(doc);
            textBox.AppendChild(para);

            // Add the text to the paragraph
            Run run = new Run(doc, StampText);
            run.Font.Size = 12; // Set the font size
            run.Font.Bold = true; // Set the font style to bold            
            para.AppendChild(run);

            // Save the modified document
            doc.Save(FileOutputPath);           
        }
    }
}
