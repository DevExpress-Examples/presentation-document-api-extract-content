using DevExpress.Docs.Presentation;
using DevExpress.Drawing;


namespace ExtractPresentationContent;

public class Program {
    public static void Main(string[] _) {
        Presentation presentation = new Presentation(File.ReadAllBytes(@"data\Sample.pptx"));

        Console.WriteLine("Choose what to extract:");
        Console.WriteLine("1. Slide text");
        Console.WriteLine("2. Text from all slides");
        Console.WriteLine("3. Text from a specific paragraph");
        Console.WriteLine("4. Slide notes");
        Console.WriteLine("5. Slide note body text");
        Console.WriteLine("6. Notes from all slides");
        Console.WriteLine("7. Slide pictures");
        Console.WriteLine("8. Pictures from all slides");

        Console.Write("Enter option number: ");
        string? input = Console.ReadLine();

        switch (input)
        {
            case "1":
                SlideText(presentation, 1);
                break;
            case "2":
                SlidesText(presentation);
                break;
            case "3":
                ParagraphText(presentation);
                break;
            case "4":
                SlideNoteText(presentation,0);
                break;
            case "5":
                SlideNoteBodyText(presentation, 0);
                break;
            case "6":
                SlidesNoteText(presentation);
                break;
            case "7":
                SaveSlidePictures(presentation, 1);
                break;
            case "8":
                SaveSlidesPictures(presentation);
                break;
            default:
                Console.WriteLine("Invalid option.");
                break;
        }
    }

    public static string SlideText(Presentation presentation, int slideNumber) {
        // Extract text from the second slide 
        string slidesText = "";

        #region Sort shapes
        var sortedShapes = presentation.Slides[slideNumber].Shapes
            .Where(shape => shape is Shape && ((Shape)shape).TextArea != null)
            .OrderBy(shape => shape.Y)
            .ThenBy(shape => shape.X);
        #endregion

        foreach (var shape in sortedShapes)
        {
            if (shape is Shape textShape)
            {
                string shapeText = textShape.TextArea.Text;

                #region Filter shapes
                if (textShape.PlaceholderSettings?.Type == PlaceholderType.SlideNumber
                    || string.IsNullOrWhiteSpace(shapeText)) continue;
                #endregion

                slidesText += shapeText + "\r\n";
            }
        }
        return slidesText;
    }

    public static string SlidesText(Presentation presentation) {
        // Extract text from all slides
        int slideNumber = 0;
        string slidesText = "";
        foreach (var slide in presentation.Slides) {
            #region Sort shapes
            var sortedShapes = slide.Shapes
                .Where(shape => shape is Shape && ((Shape)shape).TextArea != null)
                .OrderBy(shape => shape.Y)
                .ThenBy(shape => shape.X);
            #endregion

            foreach (var shape in sortedShapes) {
                if (shape is Shape textShape) {
                    string shapeText = textShape.TextArea.Text;

                    #region Filter shapes
                    if (textShape.PlaceholderSettings?.Type == PlaceholderType.SlideNumber
                        || string.IsNullOrWhiteSpace(shapeText)) continue;
                    #endregion

                    slidesText += shapeText + "\r\n";
                }
            }
            slideNumber++;
        }
        return slidesText;
    }

    public static string ParagraphText(Presentation presentation) {
        //Extract text from the specific shape's paragraph
        string paraText = "";
        Shape shape = presentation.Slides[0].Shapes.Find<Shape>(s => s.Name == "TextBox 3");
        return paraText = shape.TextArea.Paragraphs[1].Text;
    }

    public static string SlideNoteText(Presentation presentation, int slideNumber) {
        // Extract note text from the first slide 
        string slideNoteText = "";
        foreach (var noteShape in presentation.Slides[slideNumber].Notes.Shapes) {
            if (noteShape is Shape textNoteShape) {
                string noteShapeText = textNoteShape.TextArea.Text;
                #region Filter shapes
                if (textNoteShape.PlaceholderSettings?.Type == PlaceholderType.SlideNumber
                    || string.IsNullOrWhiteSpace(noteShapeText)) continue;
                #endregion
                slideNoteText += noteShapeText + "\r\n";
            }
        }
        return slideNoteText;
    }
    public static string SlideNoteBodyText(Presentation presentation, int slideNumber) {
        string notesText = "";
        foreach (var noteShape in presentation.Slides[slideNumber].Notes.Shapes) {
            if (noteShape is Shape textNoteShape
                && textNoteShape.PlaceholderSettings.Type == PlaceholderType.Body) {
                string noteShapeText = textNoteShape.TextArea.Text;
                #region Filter shapes
                if (textNoteShape.PlaceholderSettings?.Type == PlaceholderType.SlideNumber
                    || string.IsNullOrWhiteSpace(noteShapeText)) continue;
                notesText += noteShapeText + "\r\n";
                #endregion
            }
        }
        return notesText;
    }
    public static string SlidesNoteText(Presentation presentation) {
        string notesText = "";
        foreach (var slide in presentation.Slides) {
            if (slide.Notes.Shapes.Any()) {
                #region Sort shapes
                var sortedNoteShapes = slide.Notes.Shapes
                    .Where(shape => shape is Shape && ((Shape)shape).TextArea != null)
                    .OrderBy(shape => shape.Y)
                    .ThenBy(shape => shape.X);
                #endregion
                foreach (var noteShape in sortedNoteShapes) {
                    if (noteShape is Shape textNoteShape) {
                        string noteShapeText = textNoteShape.TextArea.Text;
                        #region Filter shapes
                        if (textNoteShape.PlaceholderSettings?.Type == PlaceholderType.SlideNumber
                            || string.IsNullOrWhiteSpace(noteShapeText)) continue;
                        notesText += noteShapeText + "\r\n";
                        #endregion
                    }
                }
            }
        }
        return notesText;
    }
    public static void SaveSlidePictures(Presentation presentation, int slideNumber) {
        // Extract pictures from the second slide 
        #region Sort shapes
        var sortedShapes = presentation.Slides[slideNumber].Shapes
            .Where(shape => shape is PictureShape)
            .OrderBy(shape => shape.Y)
            .ThenBy(shape => shape.X);
        #endregion

        byte imageCount = 0;
        foreach (PictureShape pictureShape in sortedShapes) {
            pictureShape.Image.Save("Slide2_Picture" + imageCount + ".png", DXImageFormat.Png);
            imageCount++;
        }

    }
    public static void SaveSlidesPictures(Presentation presentation) {
        // Extract pictures from all slides
        int slideNumber = 0;
        foreach (var slide in presentation.Slides)
        {
            #region Sort shapes
            var sortedShapes = slide.Shapes
                .Where(shape => shape is PictureShape)
                .OrderBy(shape => shape.Y)
                .ThenBy(shape => shape.X);
            #endregion

            byte imageCount = 0;
            foreach (PictureShape pictureShape in sortedShapes) {
                pictureShape.Image.Save("Slide" + slideNumber + "_Picture" + imageCount + ".png", DXImageFormat.Png);
                imageCount++;
            }
            slideNumber++;
        }
    }
}