using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

internal class PptmToPptx
{

    public static void Save(string pptmFilePath, string pptxFilePath)
    {
        // Open the .pptm file
        using (PresentationDocument pptmDoc = PresentationDocument.Open(pptmFilePath, false))
        {
            // Create a new .pptx file
            using (PresentationDocument pptxDoc = PresentationDocument.Create(pptxFilePath, PresentationDocumentType.Presentation))
            {
                // Add the necessary parts to the new .pptx file
                pptxDoc.AddPresentationPart();

                pptxDoc.PresentationPart.Presentation = new Presentation()
                {
                    SlideIdList = new SlideIdList(),
                    NotesSize = new NotesSize() { Cx = 6858000, Cy = 9144000 },
                    //SlideMasterIdList = new SlideMasterIdList()
                };
                
                //pptxDoc.PresentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId() { RelationshipId = "rId1" });

                // Clone the slide master part from the .pptm file to the new .pptx file
                SlideMasterPart newSlideMasterPart = pptmDoc.PresentationPart.SlideMasterParts.FirstOrDefault();
                pptxDoc.PresentationPart.AddPart<SlideMasterPart>(newSlideMasterPart);
                                
                UInt32 slideId = 1000;
                // Clone the presentation slides from the .pptm file to the new .pptx file
                foreach (SlidePart slidePart in pptmDoc.PresentationPart.SlideParts)
                {
                    SlidePart newSlidePart = pptxDoc.PresentationPart.AddPart<SlidePart>(slidePart);
                    //newSlidePart.Slide = (Slide)slidePart.Slide.CloneNode(true);
                    string relId = pptxDoc.PresentationPart.GetIdOfPart(newSlidePart);
                    pptxDoc.PresentationPart.Presentation.SlideIdList.Append(new SlideId() { RelationshipId = relId, Id= slideId++ });

                   
                    //SlideSize notesSlideSize = slidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree.GetFirstChild<SlideSize>();
                    //if (notesSlideSize != null)
                    //{
                    //    SlidePart newNotesSlidePart = pptxDoc.PresentationPart.GetPartById(relId) as SlidePart;
                    //    if (newNotesSlidePart.NotesSlidePart != null)
                    //    {
                    //        SlideSize newNotesSlideSize = new SlideSize() { Cx = notesSlideSize.Cx, Cy = notesSlideSize.Cy };
                    //        newNotesSlidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree.InsertAfter(newNotesSlideSize, notesSlideSize);
                    //    }
                    //}
                }

                // Save and close the new .pptx file
                pptxDoc.Save();
            }
        }
    }

   



    public static void ValidatePptxFile(string pptxFilePath)
    {
        using (PresentationDocument pptxDoc = PresentationDocument.Open(pptxFilePath, false))
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(pptxDoc);

            if (errors.Count() > 0)
            {
                foreach (ValidationErrorInfo error in errors)
                {
                    Console.WriteLine("Error: " + error.Description);
                }
            }
            else
            {
                Console.WriteLine("Validation successful. No errors found.");
            }
        }
    }


    /// <summary>
    // Selected a .pptm file (with macro storage), remove the VBA. 
    // Change the PresentationDocumentType to Presentation, and save the Presentation document with pptx extension.
    /// </summary>
    /// <param name="filePath"></param>
    public static void ConvertPptmToPptx(string filePath)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
        {
            // Look for the vbaProject part. If it is there, delete it.
            var vbaPart = presentationDocument.PresentationPart.VbaProjectPart;
            if (vbaPart is not null)
            {
                // Delete the vbaProject part and then save the document.
                presentationDocument.PresentationPart.DeletePart(vbaPart);
                presentationDocument.Save();

                // Change the document type to Presentation (i.e., remove macro-enabled features)
                presentationDocument.ChangeDocumentType(PresentationDocumentType.Presentation);
            }
        }

        
        var newfilePath = Path.ChangeExtension(filePath, ".pptx");

        // If it already exists, it will be deleted!
        if (File.Exists(newfilePath))
        {
            File.Delete(newfilePath);
        }

        // Rename the file.
        File.Move(filePath, newfilePath);
    }

    //convert pptm to pptx using openxml
    public static void ConvertPptmToPptx(string pptmFilePath, string pptxFilePath)
    {
        //This operation requires that the document be opened with ReadWrite (or Write) access
        using (Stream stream = new FileStream(pptmFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            // Open the .pptm file with ReadWrite access
            using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
            {
                // Look for the vbaProject part. If it is there, delete it.
                var vbaPart = presentationDocument.PresentationPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    // Delete the vbaProject part and then save the document.
                    presentationDocument.PresentationPart.DeletePart(vbaPart);
                    presentationDocument.Save();
                }

                    // Change the document type to Presentation (i.e., remove macro-enabled features)
                    presentationDocument.ChangeDocumentType(PresentationDocumentType.Presentation);

                // Save the document as a .pptx file
                presentationDocument.SaveAs(pptxFilePath);
            }
        }


        //using (Stream inputStream = new FileStream(pptmFilePath, FileMode.Open, FileAccess.Read))
        //{
        //    using (Stream outputStream = new FileStream(pptxFilePath, FileMode.Create, FileAccess.ReadWrite))
        //    {
        //        // Load the presentation from the input .pptm file
        //        using (PresentationDocument presentationDocument = PresentationDocument.Open(inputStream, false))
        //        {
        //            // Create a new presentation with the desired type (pptx)
        //            PresentationDocument newPresentationDocument = PresentationDocument.Create(outputStream, PresentationDocumentType.Presentation);

        //            // Clone the parts from the source document to the new document
        //            foreach (var part in presentationDocument.Parts)
        //            {
        //                newPresentationDocument.AddPart(part.OpenXmlPart, part.RelationshipId);
        //            }

        //            // Save the changes to the new .pptx file
        //            newPresentationDocument.Save();
        //        }
        //    }
        //}
    }

    public static void DuplicatePptxFile(string sourceFilePath, string newFilePath)
    {
        // Open the source PowerPoint file
        using (var sourcePresentation = PresentationDocument.Open(sourceFilePath, false))
        {
            // Create a new file to copy the content of the source file
            File.Copy(sourceFilePath, newFilePath, true);
        }
    }


}
