using Aspose.Words;
using Aspose.Words.Drawing;

namespace ConsoleApp_FileChange
{
    internal class WordDocChange
    {
        private const string FilePath = @"C:\sampleDoc\docx\Sample.docx";
        private const string FileOutputPath = @"C:\sampleDoc\docx\Sample-Output.docx";

        private const double ShapeStampWidth = 255;
        private const double ShapeStampHeight = 20;
        private const double ShapeStampXCoordinate = 50;
        private const double ShapeStampBottomPosition = 20;
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
            Shape textBox = builder.InsertShape(ShapeType.TextBox, RelativeHorizontalPosition.Page, ShapeStampXCoordinate, RelativeVerticalPosition.Page,
                top, ShapeStampWidth, ShapeStampHeight, WrapType.None);
            textBox.Stroked = false;

            //textBox.WrapSide = WrapSide.Both;
            //textBox.WrapType = WrapType.Tight;
            //textBox.DistanceTop = 0;
            //textBox.DistanceBottom = 0;
            //textBox.BehindText = true;

            //textBox.VerticalAlignment = VerticalAlignment.Top;
            textBox.TextBox.VerticalAnchor = TextBoxAnchor.Top;
            textBox.TextBox.FitShapeToText = true;
            textBox.TextBox.InternalMarginLeft = 0.0;
            textBox.TextBox.InternalMarginRight = 0.0;
            textBox.TextBox.InternalMarginBottom = 0.0;
            textBox.TextBox.InternalMarginTop = 0.0;

            // Create a paragraph inside the text box
            //textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            para.ParagraphFormat.LineUnitBefore = 0;
            para.ParagraphFormat.LineUnitAfter = 0;
            //para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            //para.ParagraphFormat.WordWrap = false;

            // Set paragraph formatting to remove spacing
            para.ParagraphFormat.SpaceBeforeAuto = false;
            para.ParagraphFormat.SpaceBefore = 0;
            para.ParagraphFormat.SpaceAfterAuto = false;
            para.ParagraphFormat.SpaceAfter = 0;

            // Add the text to the paragraph
            Run run = new Run(doc, Constants.StampText);
            run.Font.Size = 12; // Set the font size
            run.Font.Bold = true; // Set the font style to bold
            run.Font.Name = Constants.FontName;
            para.AppendChild(run);

            // Save the modified document
            doc.Save(FileOutputPath);
        }
    }
}
