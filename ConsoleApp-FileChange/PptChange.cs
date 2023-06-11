using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp_FileChange
{
    internal class PptChange
    {
        private const string presentationPath = @"C:\sampleDoc\EmployeeHandbook.pptx";
        private const string stampText = "Sanitized By Global Markets - EY Knowledge";
        internal void AddText()
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationPath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as P.SlideId).RelationshipId;

                // Get the slide part from the relationship IA.
                SlidePart slide = (SlidePart)presentationPart.GetPartById(relId);

                AddStamp(slide);
            }
        }

        private static void AddStamp(SlidePart slide)
        {
            CommonSlideData commonSlideData = slide.Slide.CommonSlideData;

            var textBody = commonSlideData.Descendants<P.TextBody>().LastOrDefault();
            if (textBody == null)
            {
                textBody = new P.TextBody();
                commonSlideData.AppendChild(textBody);
            }

            // Find the first paragraph element or create a new one if it doesn't exist
            var paragraph = textBody.Descendants<A.Paragraph>().LastOrDefault();
            if (paragraph == null)
            {
                paragraph = new A.Paragraph();
                textBody.AppendChild(paragraph);
            }

            // Create a new run and text elements for the paragraph
            var run = new A.Run();
            var text = new A.Text(stampText);
            run.AppendChild(text);


            // Get the font color from an existing text run or paragraph
            var fontColor = GetFontColorFromSlide(slide);
            var runProperties = new A.RunProperties()
            {
                FontSize = 1400,
                Bold = true
            };
            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = fontColor.Val.Value }));

            run.Append(runProperties);

            A.Paragraph paragraph1 = new A.Paragraph();
            paragraph1.AppendChild(run);

            paragraph1.ParagraphProperties = new A.ParagraphProperties()
            {
                //Level = -356616,
                //LeftMargin = 356616,
                //LineSpacing = new A.LineSpacing() { SpacingPercent = new A.SpacingPercent() { Val = Convert.ToInt32(slideHeight- veriticalSpace) } },
            };

            paragraph.InsertAfterSelf(paragraph1);
        }

        static Color GetFontColorFromSlide(SlidePart slidePart)
        {
            // Get the first text run or paragraph from the slide
            var firstTextRun = slidePart.Slide.Descendants<A.Run>().FirstOrDefault();
            var firstParagraph = slidePart.Slide.Descendants<A.Paragraph>().FirstOrDefault();

            // Retrieve the font color from the first text run or paragraph
            if (firstTextRun != null)
            {
                var solidFill = firstTextRun.Descendants<A.SolidFill>().FirstOrDefault();
                if (solidFill != null)
                {
                    var rgbColor = solidFill.Descendants<A.RgbColorModelHex>().FirstOrDefault();
                    if (rgbColor != null)
                    {
                        return new Color { Val = rgbColor.Val };
                    }
                }
            }
            else if (firstParagraph != null)
            {
                var solidFill = firstParagraph.Descendants<A.SolidFill>().FirstOrDefault();
                if (solidFill != null)
                {
                    var rgbColor = solidFill.Descendants<A.RgbColorModelHex>().FirstOrDefault();
                    if (rgbColor != null)
                    {
                        return new Color
                        {
                            Val = rgbColor.Val
                        };
                    }
                }
            }

            // Default to black if no font color is found
            return new Color { Val = "000000" };
        }


        private static void FindSlideDimension(CommonSlideData commonSlideData, SlidePart slide)
        {
            // Get the slide size information
            PresentationPart presPart = slide.GetParentParts().FirstOrDefault(z => z is PresentationPart) as PresentationPart;

            SlideSize slideSize = null;
            if (presPart != null)
            {
                slideSize = presPart.Presentation.GetFirstChild<SlideSize>();
            }

            // Get the slide width and height
            long slideWidth = slideSize.Cx.Value;
            long slideHeight = slideSize.Cy.Value;

            Console.WriteLine($"slide width: {slideWidth} and height: {slideHeight}");

            var node = commonSlideData.Descendants<A.Paragraph>().Where(x => x.InnerText.StartsWith("Over recent years"));

            //var maxSPct = commonSlideData.Descendants<A.SpacingPercent>().Max();

            var testtext = commonSlideData.Descendants<A.Paragraph>().Where(x => x.InnerText.StartsWith("Over recent years")).FirstOrDefault().InnerText;
        }

        private void CommentedCode()
        {
            //paragraph1.ParagraphProperties = new A.ParagraphProperties()
            //{
            //    //Level = -356616,
            //    //LeftMargin = 356616,
            //    //LineSpacing = new A.LineSpacing() { SpacingPercent = new A.SpacingPercent() { Val = Convert.ToInt32(slideHeight- veriticalSpace) } },

            //};

            //// Create a new text box shape 
            //A.Shape shape = new A.Shape(
            //    new A.NonVisualShapeProperties(new A.NonVisualDrawingProperties { Id = 1, Name = "TextBox Stamp" },
            //    new A.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            //    new P.ApplicationNonVisualDrawingProperties()), new A.ShapeProperties(),
            //    new A.TextBody(new BodyProperties(), paragraph1));
            //// Set the position of the text box shape 
            //shape.ShapeProperties.Transform2D = new A.Transform2D(new A.Offset { X = 5000, Y = slideHeight - 500 }
            //// Set the X and Y position in EMUs (English Metric Units) 
            ////,new A.Extents { Cx = 50000, Cy = 2000000 } 
            //);
            // Set the width and height in EMUs
            // Add the text box shape to the slide 


            // Create a new text box shape
            P.Shape textBoxShape = new P.Shape();

            //// Create non-visual shape properties for the text box shape
            //P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            //P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 2, Name = "Text Box Shape" };
            //P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties = new P.NonVisualShapeDrawingProperties();
            //nonVisualShapeProperties.Append(nonVisualDrawingProperties);
            //nonVisualShapeProperties.Append(nonVisualShapeDrawingProperties);

            //// Create visual shape properties for the text box shape
            //A.ShapeProperties shapeProperties = new A.ShapeProperties();
            //A.TextBody textBody1 = new A.TextBody();
            //A.BodyProperties bodyProperties = new A.BodyProperties();


            //// Append elements to construct the text box shape
            //bodyProperties.Append(paragraph1);
            //textBody1.Append(bodyProperties);
            //shapeProperties.Append(textBody1); 

            //// Set the positioning for the text box shape                
            //shapeProperties.Transform2D = new A.Transform2D(
            //  new A.Offset { X = 5000, Y = slideHeight - 500 } ,               
            //  new A.Extents { Cx = 50000, Cy = 2000000 }
            //);

            //// Set the non-visual and visual properties for the text box shape
            //textBoxShape.Append(nonVisualShapeProperties);
            //textBoxShape.Append(shapeProperties);

            //// Create a new shape tree if it doesn't exist
            //if (commonSlideData.ShapeTree == null)
            //    commonSlideData.ShapeTree = new P.ShapeTree();

            //// Add the new text box shape to the shape tree
            //commonSlideData.ShapeTree.AppendChild(textBoxShape);

            //// Create a new text box shape for the footer
            //// Create a new paragraph for the footer
            //var paragraph = new A.Paragraph(new A.Run(new A.Text("Footer Text")));

            //// Create paragraph properties
            //var paragraphProperties = new A.ParagraphProperties();
            ////var alignment = new A.Alignment() { Horizontal = A.TextAlignmentTypeValues.Center }; // Align the text in the center
            ////paragraphProperties.Append(alignment);

            //// Set the paragraph properties for the paragraph
            //paragraph.Append(paragraphProperties);


            //// Create a new footer if it doesn't exist
            //var footer = commonSlideData.Descendants<P.HeaderFooter>().FirstOrDefault().f;

            //// Add the new paragraph to the footer
            //footer?.InnerText="footer";
        }
    }
}
