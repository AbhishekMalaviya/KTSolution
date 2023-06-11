using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp_FileChange;

public class WordDocumentEditor
{
    public void AddText(string filePath)
    {
        using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
        {
            // Access the main document part
            MainDocumentPart mainPart = document.MainDocumentPart;

            // Get the first paragraph of the first page
            Paragraph firstParagraph = mainPart.Document.Body.Descendants<Paragraph>().FirstOrDefault();

            // Create a new run and text elements
            Run run = new Run();
            Text text = new Text("Your text here");

            // Append the text to the run
            run.Append(text);

            // Create a new paragraph with the run
            Paragraph newParagraph = new Paragraph(run);

            // Set the paragraph alignment to bottom left
            newParagraph.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Left },
                new ParagraphBorders(
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single) },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single) }
                )
            );

            // Insert the new paragraph before the existing first paragraph
            firstParagraph.InsertBeforeSelf(newParagraph);
        }
    }
}

