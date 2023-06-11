using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ConsoleApp_FileChange
{
    public class WordProcessingTextBox
    {
        public static void AddTextBoxToDocument(string filePath)
        {
            // Open the Word document using OpenXML SDK
            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // Get the main document part
                MainDocumentPart mainPart = document.MainDocumentPart;

                //// Create a new text box shape
                //var textBoxShape = new A.Shape(
                //    new A.TextBody(
                //        new A.BodyProperties(),
                //        new A.ListStyle(),
                //        new A.Paragraph(new A.Run(new A.Text("Your Text")))
                //    ),
                //    new A.ShapeProperties()
                //);

                //// Create the text box object
                //var textBox = new TextBox(textBoxShape);
                ////textBox.Style = "width:200pt;height:100pt;position:absolute;bottom:0;left:0";

                // Create a new paragraph to hold the text box
                var paragraph = new Paragraph();

                PlaceTextAtCoordinate(paragraph, "Text at 120.5,120.5", 120.1, 120.1);

                // Get the first section and add the paragraph to it
                var section = mainPart.Document.Body.Elements<SectionProperties>().First();
                section.Append(paragraph);
            }
        }

        public static void PlaceTextAtCoordinate(Paragraph para, string text, double xCoordinate, double uCoordinate)
        {
            var picRun = para.AppendChild(new Run());

            Picture picture1 = picRun.AppendChild(new Picture());

            Shapetype shapetype1 = new Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            Stroke stroke1 = new Stroke() { JoinStyle = StrokeJoinStyleValues.Miter };
            DocumentFormat.OpenXml.Vml.Path path1 = new () { AllowGradientShape = true, ConnectionPointType = ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            Shape shape1 = new Shape() { Id = "Text Box 2", Style = string.Format("position:absolute;margin-left:{0:F1}pt;margin-top:{1:F1}pt;width:187.1pt;height:29.7pt;z-index:251657216;visibility:visible;mso-wrap-style:square;mso-width-percent:400;mso-height-percent:200;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text;mso-width-percent:400;mso-height-percent:200;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", xCoordinate, uCoordinate), Stroked = false };

            TextBox textBox1 = new TextBox() { Style = "mso-fit-shape-to-text:t" };

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph2 = new Paragraph();

            Run run2 = new Run();
            Text text2 = new Text();
            text2.Text = text;

            run2.Append(text2);

            paragraph2.Append(run2);

            textBoxContent1.Append(paragraph2);

            textBox1.Append(textBoxContent1);
            TextWrap textWrap1 = new TextWrap() { Type = WrapValues.Square };

            shape1.Append(textBox1);
            shape1.Append(textWrap1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);
        }

        

    }
}
