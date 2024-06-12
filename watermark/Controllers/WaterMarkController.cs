using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using Aspose.Imaging;
using Aspose.Imaging.Brushes;
using Aspose.Imaging.FileFormats.Jpeg;
using Aspose.Imaging.ImageOptions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.ThreeD;
using Aspose.Slides;
using Aspose.Slides.Export;
using License = Aspose.Slides.License;
using ShapeType = Aspose.Slides.ShapeType;
using FillType = Aspose.Slides.FillType;
using SaveFormat = Aspose.Slides.Export.SaveFormat;
using SeekOrigin = System.IO.SeekOrigin;
using Aspose.Pdf.Text;
using Aspose.Pdf;
using Document = Aspose.Words.Document;
using watermark.Services;
using Aspose.PSD.FileFormats.Psd.Layers.LayerResources;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Operators;
using HorizontalAlignment = Aspose.Pdf.HorizontalAlignment;
using VerticalAlignment = Aspose.Pdf.VerticalAlignment;
using Aspose.Words.Pdf2Word.FixedFormats;

namespace watermark.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WaterMarkController : ControllerBase
    {
        private readonly ILogger<WaterMarkController> _logger;

        public WaterMarkController(ILogger<WaterMarkController> logger)
        {
            _logger = logger;

        }

        [HttpPost()]
        [Route("ImageWatermarksimple")]
        public async Task<ActionResult> ImageAsposeSimple(IFormFile file)
        {
            string templatesFolder = @"C:\\aspose\\img\\";
            string dataDir = templatesFolder;
            if (file == null || file.Length == 0)
            {
                return BadRequest();
            }
            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);
                using (Aspose.Imaging.Image image = Aspose.Imaging.Image.Load(ms))
                {
                    string theString = "Avepoint VietNam";

                    Graphics graphics = new Graphics(image);
                    SizeF sz = graphics.Image.Size;
                    Aspose.Imaging.Font font = new Aspose.Imaging.Font("Times New Roman", 30, Aspose.Imaging.FontStyle.Regular);


                    SolidBrush brush = new SolidBrush(Aspose.Imaging.Color.Transparent);
                    //brush.Opacity = 1;

                    StringFormat format = new StringFormat();
                    format.Alignment = StringAlignment.Center;
                    format.FormatFlags = StringFormatFlags.DirectionVertical;

                    graphics.DrawString(theString, font, brush, 100, 100, format);
                    image.Save("C:\\aspose\\testWatermark\\" + "resultsimple.jpg");
                }
            }
            return Ok();
        }

        [HttpPost()]
        [Route("ImageWatermarksimplecheck")]
        public async Task<string> ImageAsposeSimplecheck(IFormFile file)
        {
            string templatesFolder = @"C:\\aspose\\img\\logo ave\\logo4.png";
            string dataDir = templatesFolder;
            if (file == null || file.Length == 0)
            {
                return "";
            }
            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);
                using (Aspose.Imaging.Image image = Aspose.Imaging.Image.Load(ms))
                {
                    string theString = "Avepoint VietNam";
                    Aspose.Imaging.Graphics graphics = new Aspose.Imaging.Graphics(image);
                    graphics.DrawImage(Aspose.Imaging.Image.Load(templatesFolder) , 1, 1);
                    var a = graphics.Image;
                }
                return "";
            }
        }

        //[HttpPost()]
        //[Route("AddWaterMark")]
        //public async Task<ActionResult> AddWaterMark(IFormFile file)
        //{
        //    var name = Path.GetFileName(file.FileName);
        //    string f = Path.GetFileNameWithoutExtension(file.FileName) + ".png";
        //    if(name != null)
        //    {
        //        var stream = new MemoryStream();
        //        await file.CopyToAsync(stream);
        //        string text = "lucas.dinh";
        //        //using(var stream = new MemoryStream())
        //        //{
        //        //    await file.CopyToAsync(stream);
        //        //    var a = stream.ToArray();
        //        //    return File(a,"image/png",f);
        //        //}
        //        using(Bitmap bitmap = new Bitmap(stream, false))
        //        {
        //            using(Graphics graphics = Graphics.FromImage(bitmap))
        //            {
        //                Brush brush = new SolidBrush(Color.FromArgb(90,Color.Gray));
        //                Font font = new Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point);
        //                SizeF textSize = new SizeF();
        //                textSize = graphics.MeasureString(text, font);
        //                for(int i = 1; i <= 10; i++)
        //                {
        //                    for(int j = 1; j <= 10; j++)
        //                    {
        //                        Point position = new Point(bitmap.Width*i/10, bitmap.Height*j/10);
        //                        graphics.DrawString(text, font, brush, position);
        //                    }
        //                }
        //                using(MemoryStream ms = new MemoryStream())
        //                {
        //                    bitmap.Save(ms,ImageFormat.Png);
        //                    ms.Position = 0;
        //                    return File(ms.ToArray(), "image/png", f);
        //                }
        //            }
        //        }
        //    }
        //    return Ok();
        //}


        //[HttpPost()]
        //[Route("TestAspose")]
        //public async Task<ActionResult> TestAsposeWord(IFormFile file)
        //{
        //    var stream = new MemoryStream();
        //    await file.CopyToAsync(stream);
        //    Document docA = new Document(stream);

        //    // Inisialize a DocumentBuilder
        //    DocumentBuilder builder = new DocumentBuilder(docA);

        //    // Insert text to the document A start
        //    docA.Save("C:\\aspose\\output_AB2.png");
        //    //builder.MoveToDocumentStart();

        //    builder.Write("First Hello World paragraph");

        //    // Open an existing document B
        //    Document docB = new Document("C:\\aspose\\documentB.docx");

        //    // Add document B to the and of document A, preserving document B formatting
        //    docA.AppendDocument(docB, ImportFormatMode.KeepSourceFormatting);

        //    // Save the output as PDF
        //    docA.Save("C:\\aspose\\output_AB.png");
        //    return Ok();
        //}

        //[HttpPost()]
        //[Route("watermarkAspose")]
        //public async Task<ActionResult> WatermarkAspose(IFormFile file)
        //{
        //    var stream = new MemoryStream();
        //    await file.CopyToAsync(stream);
        //    Document doc = new Document(stream);

        //    // Add a plain text watermark.
        //    doc.Watermark.SetText("lucas.dinh");

        //    // If we wish to edit the text formatting using it as a watermark,
        //    // we can do so by passing a TextWatermarkOptions object when creating the watermark.
        //    TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
        //    textWatermarkOptions.FontFamily = "Arial";
        //    textWatermarkOptions.FontSize = 36;
        //    textWatermarkOptions.Color = Color.Red;
        //    textWatermarkOptions.Layout = WatermarkLayout.Diagonal;
        //    textWatermarkOptions.IsSemitrasparent = false;

        //    doc.Watermark.SetText("lucas.dinh", textWatermarkOptions);

        //    doc.Save("C:\\aspose\\TextWatermark.docx");

        //    // We can remove a watermark from a document like this.
        //    if (doc.Watermark.Type == WatermarkType.Text)
        //        doc.Watermark.Remove();

        //    return Ok();
        //}

        [HttpPost()]
        [Route("ImageWatermark/png-jpg-jpeg-webp")]
        public async Task<ActionResult> ImageAspose(IFormFile file)
        {
            string templatesFolder = @"C:\\aspose\\img\\";
            string dataDir = templatesFolder;
            if(file == null || file.Length == 0)
            {
                return BadRequest();
            }

            //var fs = new FileStream(dataDir + "logo.webp", FileMode.Open);
            //await file.CopyToAsync(fs);

            //var path1 = Path.GetFullPath(file.FileName);
            //var path = Path.Combine(Directory.GetCurrentDirectory(), file.FileName);
            // Load an existing JPG image
            //using (Image image = Image.Load(dataDir + "logo.webp"))
            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);
                //Image img = Image.Load(ms);
                using (Aspose.Imaging.Image image = Aspose.Imaging.Image.Load(ms))
                {
                    // Declare a String object with Watermark Text
                    //string theString = "45 Degree Rotated Text";
                    string theString = "Avepoint VietNam";


                    // Create and initialize an instance of Graphics class and Initialize an object of SizeF to store image Size
                    Graphics graphics = new Graphics(image);
                    SizeF sz = graphics.Image.Size;

                    // Creates an instance of Font, initialize it with Font Face, Size and Style
                    Aspose.Imaging.Font font = new Aspose.Imaging.Font("Times New Roman", 30, Aspose.Imaging.FontStyle.Bold);
                    

                    // Create an instance of SolidBrush and set its various properties
                    SolidBrush brush = new SolidBrush(Aspose.Imaging.Color.FromArgb(120, Aspose.Imaging.Color.IndianRed));
                    //brush.Color = Color.Red;
                    brush.Opacity = 1;

                    // Initialize an object of StringFormat class and set its various properties
                    StringFormat format = new StringFormat();
                    format.Alignment = StringAlignment.Far;
                    format.FormatFlags = StringFormatFlags.DirectionVertical;

                    // Create an object of Matrix class for transformation
                    Aspose.Imaging.Matrix matrix = new Aspose.Imaging.Matrix();
                    //for(int i = 0; i < 5; i++)
                    //{
                    //for(int j = 0; j < 5; j++)
                    //{
                    //    // First a translation then a rotation                
                    //        matrix.Translate(sz.Width*j/5, sz.Height*4/5);
                    //    if (i ==0 & j == 0)
                    //    {
                    //        matrix.Rotate(-45.0f);
                    //    }

                    //    // Set the Transformation through Matrix
                    //    graphics.Transform = matrix;

                    //    // Draw the string on Image Save output to disk
                    //    graphics.DrawString(theString, font, brush,0 ,0 , format);
                    //}
                    //}
                    // First a translation then a rotation

                    matrix.Rotate(-45.0f);
                    for (int i = 0; i <= 10; i++)
                    {
                        for (int j = 0; j <= 10; j++)
                        {
                            //matrix.Translate(sz.Width* i/ 20, sz.Height* i/ 20);

                            // Set the Transformation through Matrix
                            graphics.Transform = matrix;

                            // Draw the string on Image Save output to disk
                            graphics.DrawString(theString, font, brush, -sz.Height + (sz.Width + sz.Height) * i / 10, (sz.Height + sz.Width) * j / 10, format);

                        }

                    }
                    image.Save("C:\\aspose\\testWatermark\\" + "result.jpg");
                }
            }
            //System.IO.File.Delete(dataDir + "result.jpg");
            return Ok();
        }


        // add word
        [HttpPost()]
        [Route("addwatermmarkWord")]
        public async Task<ActionResult> AddWatermarkWord(IFormFile file)
        {
            var docDir = @"C:\\aspose\\document\\";
            var imgDir = @"C:\\aspose\img\\logo ave\\";
            
            MemoryStream ms = new MemoryStream();
            //await file.CopyToAsync(ms);
            //var docTest = new Aspose.Words.Document(ms);
            //Aspose.Words.Document doc = new Aspose.Words.Document(ms);
            //ImageWatermarkOptions options = new ImageWatermarkOptions()
            //{
            //    Scale = 0,
            //    IsWashout = true,
            //};
            //doc.Watermark.SetImage((imgDir + "logo4.png"), options);
            //doc.Save("C:\\aspose\\testWatermark\\" + "docAddWaterImg.docx");

            // add text watermark
            Aspose.Words.Document docText = new Aspose.Words.Document(ms);
            TextWatermarkOptions optionsText = new TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = System.Drawing.Color.Transparent,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true
            };
            docText.Watermark.SetText("AVEPOINT VIETNAM", optionsText);
            docText.Save("C:\\aspose\\testWatermark\\" + "AddTextWatermark_out.docx");

            return Ok();
        }

        // add and verify excel
        [HttpPost()]
        [Route("excel")]
        public async Task<string> Excel (IFormFile file)
        {
            string FilePath = @"C:\\aspose\\document\\";

            string FileName = FilePath + "MSF HCIS_SHIP_HATS.xlsx";
            MemoryStream stream = new MemoryStream();
            await file.CopyToAsync(stream);

            //Instantiate a new Workbook

            Workbook workbook = new Workbook(stream);

            //Get the first default sheet
            foreach( Worksheet sheet in workbook.Worksheets)
            {
                ////Aspose.Cells.Drawing.Shape wordart1 = sheet.Shapes.AddWordArt(PresetWordArtStyle.WordArtStyle1, "avepoint vietnam" , 5, 5, 5, 5, 5, 1);
                //Aspose.Cells.Drawing.Shape wordart = sheet.Shapes.AddTextEffect(MsoPresetTextEffect.TextEffect2,
                //                            "", "Arial Black", 50, false, true
                //                            , 18, 8, 1, 1, 130, 800);

                //// Get the fill format of the word art
                //Aspose.Cells.Drawing.FillFormat wordArtFormat = wordart.Fill;
                //wordart.Name = "okok";

                //// Set the transparency
                //wordArtFormat.Transparency = 1.0;

                //// Make the line invisible
                //Aspose.Cells.Drawing.LineFormat lineFormat = wordart.Line;
                var a = sheet.Shapes.FindIndex(a => a.Name == "avepoint");
                if(a != -1)
                {
                    return "exist watermark in excel";
                }
                else
                {
                    return "not exist watermark in excel";
                }
            }

            workbook.Save("C:\\aspose\\testWatermark\\" + "MSF HCIS_SHIP_HATS2.xlsx");
            return "";
            //return Ok();
        }

        [HttpPost()]
        [Route("powerpoint_ppt")]
        public async Task<ActionResult> Powerpoint(IFormFile file)
        {
            string PathForWatermarkPptFile = @"C:\\aspose\\document\\";
            string imgPath = @"C:\\aspose\\img\\logo ave\\";
            MemoryStream ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Seek(0, SeekOrigin.Begin);
            Presentation WatermarkPptxPresentation = new Presentation(ms);

            foreach (IMasterSlide masterSlide in WatermarkPptxPresentation.Masters)
            {
                IAutoShape PptxWatermark = masterSlide.Shapes.AddAutoShape(ShapeType.NotDefined,
                    WatermarkPptxPresentation.SlideSize.Size.Width - 120,
                    WatermarkPptxPresentation.SlideSize.Size.Height - 30,
                    120, 10);

                PptxWatermark.FillFormat.FillType = FillType.NoFill;

                //Adding Text frame with watermark text
                ITextFrame WatermarkText = PptxWatermark.AddTextFrame("AVEPOINT VIETNAM");

                //Setting textual properties of the watermark text
                IPortionFormat WatermarkTextFormat = WatermarkText.Paragraphs[0].Portions[0].PortionFormat;
                WatermarkTextFormat.FontBold = NullableBool.True;
                WatermarkTextFormat.FontItalic = NullableBool.True;
                WatermarkTextFormat.FontHeight = 10;
                WatermarkTextFormat.FillFormat.FillType = FillType.Solid;
                WatermarkTextFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.FromArgb(200, System.Drawing.Color.Red);

                //Locking Pptx watermark shape to be uneditable in PowerPoint 
                PptxWatermark.AutoShapeLock.TextLocked = false;
                PptxWatermark.AutoShapeLock.SelectLocked = false;
                PptxWatermark.AutoShapeLock.PositionLocked = false;

            }
            WatermarkPptxPresentation.Save("C:\\aspose\\testWatermark\\" + "WatermarkPresentation.pptx",
                SaveFormat.Pptx);


            Presentation WatermarkPptxPresentation1 = new Presentation(ms);
            System.Drawing.Image WatermarkLogo = (System.Drawing.Image)new System.Drawing.Bitmap(imgPath + "logo4.png");
            IPPImage WatermarkImage = WatermarkPptxPresentation1.Images.AddImage(WatermarkLogo);

            //Accessing the master slides for adding watermark image
            foreach (IMasterSlide masterSlide in WatermarkPptxPresentation1.Masters)
            {
                //Adding a Ppt watermark shape for logo image
                IPictureFrame PptxWatermark = masterSlide.Shapes.AddPictureFrame(ShapeType.Rectangle, WatermarkPptxPresentation1.SlideSize.Size.Width - 150,
                    WatermarkPptxPresentation1.SlideSize.Size.Height - 60,
                    120, 80, WatermarkImage);

                //Set the rotation angle of the shape
                //PptxWatermark.Rotation = 325;

                //Lock Pptx watermark image shape for protection in PowerPoint 
                PptxWatermark.ShapeLock.SizeLocked = true;
                PptxWatermark.ShapeLock.SelectLocked = true;
                PptxWatermark.ShapeLock.PositionLocked = true;
            }

            //Saving the image watermark PPTX  presentation file
            WatermarkPptxPresentation1.Save("C:\\aspose\\testWatermark\\" + "ImageWatermarkedPresentation.pptx",
                SaveFormat.Pptx);

            //System.IO.File.Delete("C:\\aspose\\testWatermark\\" + "WatermarkPresentation.pptx");
            //System.IO.File.Delete("C:\\aspose\\testWatermark\\" + "ImageWatermarkedPresentation.pptx");
            return Ok();

        }

        // add ppt
        [HttpPost()]
        [Route("powerpoint_pptAdd")]
        public async Task<string> PowerpointPPT(IFormFile file)
        {
            MemoryStream ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Seek(0, SeekOrigin.Begin);
            using Presentation WatermarkPptxPresentation = new Presentation(ms);
            {
                ISlide slide = WatermarkPptxPresentation.Slides[0];
                IAutoShape watermarkShape = slide.Shapes.AddAutoShape(ShapeType.Triangle, 0, 0, 0, 0);
                watermarkShape.Name = "Avepoint VietNam";
                WatermarkPptxPresentation.Save("C:\\aspose\\testWatermark\\" + "ImageWatermarkedPresentation1111.pptx",
                SaveFormat.Pptx);
            }
            return "";
        }


        // verify ppt
        [HttpPost()]
        [Route("powerpoint_pptveirfy")]
        public async Task<string> PowerpointVerify(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName);
            MemoryStream ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Seek(0, SeekOrigin.Begin);
            using Presentation WatermarkPptxPresentation = new Presentation(ms);
            {
                ISlide slide = WatermarkPptxPresentation.Slides[0];
                if(slide.Shapes.ToList().FindIndex(x => x.Name == "avepoint") != -1)
                {
                    return "exist watermark";
                }
                else
                {
                    return "not exist";
                }
            }
            return "";

        }

        [HttpPost()]
        [Route("pdf")]
        public async Task<ActionResult> Pdf(IFormFile file, IFormFile fileLogo)
        {
            string dataDir = @"C:\\aspose\\document\\";
            string imgDir = @"C:\\aspose\\img\\logo ave\\";
            //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(dataDir + "CI-CD-Basics.pdf");
            using (MemoryStream ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);

                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ms);

                // Create text stamp
                TextStamp textStamp = new TextStamp("AVEPOINT VIETNAM");
                // Set origin
                var heightPage = pdfDocument.PageInfo.Height;
                var widthPage = pdfDocument.PageInfo.Width;
                textStamp.XIndent = heightPage - 10;
                textStamp.YIndent = widthPage - 80;

                // Set text properties
                textStamp.TextState.Font = FontRepository.FindFont("Arial");
                textStamp.TextState.FontSize = 12;
                textStamp.TextState.FontStyle = FontStyles.Italic;
                textStamp.TextState.ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Red);
                textStamp.Opacity = 50;

                // Set Stamp ID for text watermark to later identify it
                textStamp.setStampId(123456);

                // Add stamp to particular page
                foreach (var page in pdfDocument.Pages)
                {
                    page.AddStamp(textStamp);
                    //var ab = new System.Drawing.Bitmap(1,1);
                    ////ab.MakeTransparent();
                    //System.Drawing.Image img = ab;
                    ////System.Drawing.Image img = new System.Drawing.Bitmap(imgDir + "logo4.png");
                    //page.Watermark = new Aspose.Pdf.Watermark(img);
                }
                //var dataDirNew = dataDir + "Add_Text_Watermark.pdf";
                // Save output document
                pdfDocument.Save("C:\\aspose\\testWatermark\\" + "Add_Text_Watermark.pdf");

                //System.IO.File.Delete(dataDirNew);
                //return Ok();
            }

            Aspose.Pdf.Document document = new Aspose.Pdf.Document("C:\\aspose\\testWatermark\\" + "Add_Text_Watermark.pdf");
            Annotation item = document.Pages[1].Annotations[0];
            if (item is StampAnnotation annot)
            {
                TextAbsorber ta = new TextAbsorber();
                XForm ap = annot.Appearance["AVEPOINT VIETNAM"];
                ta.Visit(ap);
                Console.WriteLine(ta.Text);
            }



            using (MemoryStream ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);

                Aspose.Pdf.Document pdf = new Aspose.Pdf.Document(ms);
                MemoryStream mslogo = new MemoryStream();
                await fileLogo.CopyToAsync(mslogo);
                mslogo.Seek(0, SeekOrigin.Begin);
                ImageStamp imgStamp = new ImageStamp(mslogo);
                //ImageStamp imgStamp = new ImageStamp(imgDir + "logo1.png");
                //var height = imgStamp.Height;
                //var width = imgStamp.Width;
                imgStamp.Background = false;
                imgStamp.Height = 50;
                imgStamp.Width = 100;
                //imgStamp.TopMargin = 200;
                //imgStamp.RightMargin = 200;
                imgStamp.XIndent = pdf.PageInfo.Height ;   
                imgStamp.YIndent = pdf.PageInfo.Width - 120;
                imgStamp.Opacity = 0.5;

                imgStamp.setStampId(12345678);
                foreach( Page page in pdf.Pages)
                {
                    page.AddStamp(imgStamp);
                }

                pdf.Save("C:\\aspose\\testWatermark\\" + "imgPdf.pdf");
            }
            //System.IO.File.Delete("C:\\aspose\\testWatermark\\" + "Add_Text_Watermark.pdf");
            //System.IO.File.Delete("C:\\aspose\\testWatermark\\" + "imgPdf.pdf");

            
            //var scene = Scene.FromFile("C:\\aspose\\document\\" + "Add_Text_Watermark.pdf");
            //scene.Save("C:\\aspose\\testWatermark\\" + "Output.3ds");
            return Ok();
        }



        // add to pdf
        [HttpPost()]
        [Route("pdfverify")]
        public async Task<ActionResult> PdfVerify(IFormFile file)
        {
            string dataDir = @"C:\\aspose\\document\\";
            string imgDir = @"C:\\aspose\\img\\logo ave\\";
            using (MemoryStream ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);

                Aspose.Pdf.Document doc = new Aspose.Pdf.Document(ms);
                string test = "";
                string atext = test.Trim();
                var atextc = test.Split("\n");
                int leng = test.Split("\n").Length;
                WatermarkArtifact artifact = new WatermarkArtifact();
                artifact.SetText(new FormattedText("avepoint",System.Drawing.Color.Red, Aspose.Pdf.Facades.FontStyle.Courier, EncodingType.Identity_h, true, 72));
                artifact.ArtifactHorizontalAlignment = HorizontalAlignment.Center;
                artifact.ArtifactVerticalAlignment = VerticalAlignment.Center;
                artifact.Rotation = 45;
                artifact.Opacity = 0.5;
                artifact.Subtype = Artifact.ArtifactSubtype.Watermark;
                artifact.IsBackground = false;
                doc.Pages[1].Artifacts.Add(artifact);
                doc.Save("C:\\aspose\\testWatermark\\" + "Add_Text_Watermark11.pdf");



            }

            return Ok();
        }


        // verify pdf
        [HttpPost()]
        [Route("pdfverifypage")]
        public async Task<string> Pdfverifypage(IFormFile file)
        {
            string dataDir = @"C:\\aspose\\document\\";
            string imgDir = @"C:\\aspose\\img\\logo ave\\";
            using (MemoryStream ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);

                Aspose.Pdf.Document doc = new Aspose.Pdf.Document(ms);

                foreach (Artifact artifact in doc.Pages[1].Artifacts)
                {
                    // If artifact type is watermark, increate the counter
                    if (artifact.Subtype == Artifact.ArtifactSubtype.Watermark)
                    {
                        return "exist watermark";
                    }
                }
                return "not exist";



            }

            return "";
        }


        // verify word
        [HttpPost()]
        [Route("Verification")]
        public async Task<string> Verification(IFormFile file)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);
                Document doc = new Document(ms);

                if (doc.Watermark.Type == WatermarkType.Text || doc.Watermark.Type == WatermarkType.Image)
                {
                    //doc.Watermark.Remove();
                    return "exist watermark in this document";
                }
                if (doc.Watermark.Type == WatermarkType.None)
                {
                    return "not exist watermark in this document";
                }

                doc.Save("C:\\aspose\\verificationWatermark\\" + file.FileName);
                //Returns null if this file has no watermark,If there is a watermark, return the watermark content
            }
            return "error verification";
        }

    }
}
