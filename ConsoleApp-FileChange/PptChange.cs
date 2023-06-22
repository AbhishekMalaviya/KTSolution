using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ConsoleApp_FileChange
{
    internal class PptChange
    {
        private const string PresentationPath = @"C:\sampleDoc\ppt\EmployeeHandbook.pptx";
        private const string StampText = "Sanitized by Global Markets – EY Knowledge";
        private const long ShapeStampWidth = 4132413;
        private const long ShapeStampHeight = 220060;
        private const long ShapeStampXCoordinate = 127000;
        private const long ShapeStampYCoordinate = 6438900;


        internal static void AddText()
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(PresentationPath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as P.SlideId).RelationshipId;

                // Get the slide part from the relationship IA.
                SlidePart slide = (SlidePart)presentationPart.GetPartById(relId);

                AddStampToFixPosition(slide);
            }
        }

        private static void AddStampToFixPosition(SlidePart slide)
        {
            long slideHeight = FindSlideSize(slide).Cy.Value;

            CommonSlideData commonSlideData = slide.Slide.CommonSlideData;
            var shapeTree = commonSlideData.Descendants<P.ShapeTree>().FirstOrDefault();

            var shape = new DocumentFormat.OpenXml.Presentation.Shape();
            BindSlideProperties(slideHeight, shapeTree, shape);
            
            P.TextBody textBody;
            Paragraph paragraph;
            BindParagraphAndTextBodyProperties(out textBody, out paragraph);

            A.Run run = new A.Run();
            RunProperties runProperties = BindRunProperties();

            A.Text text = new A.Text(StampText);

            run.Append(runProperties);
            run.Append(text);
            paragraph.Append(run);
            textBody.Append(paragraph);
            shape.Append(textBody);
            shapeTree.Append(shape);
        }

        private static SlideSize FindSlideSize(SlidePart slide)
        {
            // Get the slide size information
            PresentationPart presPart = slide.GetParentParts().FirstOrDefault(z => z is PresentationPart) as PresentationPart;

            SlideSize slideSize = null;
            if (presPart != null)
            {
                slideSize = presPart.Presentation.GetFirstChild<SlideSize>();
            }
            return slideSize;
        }

        private static void BindSlideProperties(long slideHeight, ShapeTree? shapeTree, P.Shape shape)
        {
            shape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties();
            shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
            {
                Name = "Shape Stamp",
                Id = (UInt32)shapeTree.ChildElements.Count - 1
            });
            shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties());
            shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

            shape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
            shape.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
            shape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
            {
                X = ShapeStampXCoordinate,
                Y = slideHeight - ShapeStampHeight - 100000,
            });
            shape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
            {
                Cx = ShapeStampWidth,
                Cy = ShapeStampHeight,
            });
            shape.ShapeProperties.Append(new PresetGeometry
            {
                Preset = ShapeTypeValues.Rectangle
            });
        }

        private static void BindParagraphAndTextBodyProperties(out P.TextBody textBody, out Paragraph paragraph)
        {
            textBody = new P.TextBody();
            A.BodyProperties bodyProps = new A.BodyProperties();
            bodyProps.Wrap = A.TextWrappingValues.None;
            bodyProps.Vertical = A.TextVerticalValues.Horizontal;
            textBody.Append(bodyProps);

            paragraph = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties()
            {
                Alignment = TextAlignmentTypeValues.Left
            };
            paragraph.Append(paragraphProperties1);
        }        

        private static RunProperties BindRunProperties()
        {
            var runProperties = new A.RunProperties()
            {
                FontSize = 1400,
                Bold = true,
                Language = "en-US",
                Dirty = false,
                SmartTagClean = false
            };
            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = "FFFF00" })); // Set font color to white
            runProperties.Append(new A.LatinFont()
            {
                Typeface = "EYInterstate",
                CharacterSet = 0,
                PitchFamily = 2,
                Panose = "02000503020000020004"
            });
            return runProperties;
        }

        

        private static void AddStamp(SlidePart slide)
        {
            CommonSlideData commonSlideData = slide.Slide.CommonSlideData;

            var textBody = commonSlideData.Descendants<P.TextBody>().LastOrDefault();
            if (textBody == null)
            {
                textBody = new P.TextBody();
                commonSlideData.Append(textBody);
            }

            // Find the first paragraph element or create a new one if it doesn't exist
            var paragraph = textBody.Descendants<A.Paragraph>().LastOrDefault();
            if (paragraph == null)
            {
                paragraph = new A.Paragraph();
                textBody.Append(paragraph);
            }

            // Create a new run and text elements for the paragraph
            var run = new A.Run();
            var text = new A.Text(StampText);
            run.Append(text);


            // Get the font color from an existing text run or paragraph
            //var fontColor = GetFontColorFromSlide(slide);
            var runProperties = new A.RunProperties()
            {
                FontSize = 1400,
                Bold = true
            };
            //runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = fontColor.Val.Value }));

            run.Append(runProperties);

            A.Paragraph paragraph1 = new A.Paragraph();
            paragraph1.Append(run);

            paragraph1.ParagraphProperties = new A.ParagraphProperties()
            {
                //Level = -356616,
                //LeftMargin = 356616,
                //LineSpacing = new A.LineSpacing() { SpacingPercent = new A.SpacingPercent() { Val = Convert.ToInt32(slideHeight- veriticalSpace) } },
            };

            paragraph.InsertAfterSelf(paragraph1);
        }

        public static void AddShape()
        {
            using (var presentation = PresentationDocument.Open(PresentationPath, true))
            {
                var tree = presentation
                    .PresentationPart
                    .SlideParts
                    .ElementAt(0)
                    .Slide
                    .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                    .First();

                var shape = new DocumentFormat.OpenXml.Presentation.Shape();

                shape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties();
                shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                {
                    Name = "My Shape",
                    Id = (UInt32)tree.ChildElements.Count - 1
                });
                shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties());
                shape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

                shape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
                shape.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
                shape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
                {
                    X = ShapeStampXCoordinate,
                    Y = ShapeStampYCoordinate,
                });
                shape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
                {
                    Cx = ShapeStampWidth,
                    Cy = ShapeStampHeight,
                });
                shape.ShapeProperties.Append(new PresetGeometry
                {
                    Preset = ShapeTypeValues.Rectangle
                });
                shape.ShapeProperties.Append(new SolidFill
                {
                    SchemeColor = new SchemeColor
                    {
                        Val = SchemeColorValues.Accent2
                    }
                });
                shape.ShapeProperties.Append(new A.Outline(new NoFill()));

                //P.TextBody textBody = new P.TextBody();

                //A.Paragraph paragraph1 = new A.Paragraph();
                //A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties()
                //{
                //    Alignment = TextAlignmentTypeValues.Left
                //};
                //paragraph1.Append(paragraphProperties1);

                //textBody.Append(paragraph1);

                //A.Run run = new A.Run();
                //var text = new A.Text(StampText);
                //run.Append(text);
                //var runProperties = new A.RunProperties()
                //{
                //    FontSize = 1400,
                //    Bold = true,
                //};

                //run.Append(runProperties);

                //paragraph1.Append(run);


                //shape.Append(textBody);

                tree.Append(shape);
            }
        }


        //static Color GetFontColorFromSlide(SlidePart slidePart)
        //{
        //    // Get the first text run or paragraph from the slide
        //    var firstTextRun = slidePart.Slide.Descendants<A.Run>().FirstOrDefault();
        //    var firstParagraph = slidePart.Slide.Descendants<A.Paragraph>().FirstOrDefault();

        //    // Retrieve the font color from the first text run or paragraph
        //    if (firstTextRun != null)
        //    {
        //        var solidFill = firstTextRun.Descendants<A.SolidFill>().FirstOrDefault();
        //        if (solidFill != null)
        //        {
        //            var rgbColor = solidFill.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        //            if (rgbColor != null)
        //            {
        //                return new Color { Val = rgbColor.Val };
        //            }
        //        }
        //    }
        //    else if (firstParagraph != null)
        //    {
        //        var solidFill = firstParagraph.Descendants<A.SolidFill>().FirstOrDefault();
        //        if (solidFill != null)
        //        {
        //            var rgbColor = solidFill.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        //            if (rgbColor != null)
        //            {
        //                return new Color
        //                {
        //                    Val = rgbColor.Val
        //                };
        //            }
        //        }
        //    }

        //    // Default to black if no font color is found
        //    return new Color { Val = "000000" };
        //}


        

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
            //commonSlideData.ShapeTree.Append(textBoxShape);

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
