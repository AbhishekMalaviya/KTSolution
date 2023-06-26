using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace ConsoleApp_FileChange
{
    internal class PdfChange
    {
        private const string FilePath = @"C:\sampleDoc\pdf\The-Smarter-Store-Window.pdf";
        private const string FileOutputPath = @"C:\sampleDoc\pdf\The-Smarter-Store-Window-Output.pdf";
        
        
        internal static void AddStamp()
        {  
            // Open document
            Document pdfDocument = new Document(FilePath);

            // Get particular page
            Page pdfPage = pdfDocument.Pages[1];

            // Create text fragment
            TextFragment textFragment = new TextFragment(Constants.StampText);
            textFragment.Position = new Position(10, 10);

            // Set text properties
            textFragment.TextState.FontSize = 14;
            textFragment.TextState.Font = FontRepository.FindFont("Courier New", FontStyles.Bold);
            
            // Create TextBuilder object
            TextBuilder textBuilder = new TextBuilder(pdfPage);

            // Append the text fragment to the PDF page
            textBuilder.AppendText(textFragment);
                        
            // Save resulting PDF document.
            pdfDocument.Save(FileOutputPath);
        }
    }
}
