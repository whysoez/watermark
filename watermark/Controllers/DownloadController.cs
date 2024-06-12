using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;
using SkiaSharp;
using Aspose.Cells.Charts;
using Aspose.Words.Layout;

namespace watermark.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DownloadController : Controller
    {
        [HttpGet()]
        [Route("handover")]
        public async Task<ActionResult> Handover()
        {
            double marginLeft = 24;
            double marginRight = 24;
            double topStatic = 24;

            //var crmInfo = new CrmDataByDocTypeViewModel()
            //{
            //    recordId = property.Id.ToString(),
            //    entity = "hcis_Property",
            //};
            string header = "Handover acknowledgement";
            string escort = "Escort Statement";
            string escortData = "I acknowledge that I have received Emma from the home for Outing on date time as per record.";
            string escortSign = "Escort Person Signature";

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            double width = 1000;
            double height = 672;
            builder.PageSetup.PageWidth = width;
            builder.PageSetup.PageHeight = height;

            // Create style header
            Style style = doc.Styles.Add(StyleType.Paragraph, "Header");
            style.Font.Size = 18;
            style.Font.Name = "Open Sans";
            style.Font.Bold = true;
            style.Font.Color = Color.FromArgb(255, 255, 255);

            // custom style title
            Style styleTitle = doc.Styles.Add(StyleType.Paragraph, "Title");
            styleTitle.Font.Size = 14;
            styleTitle.Font.Name = "Open Sans";
            styleTitle.Font.Color = Color.FromArgb(85, 85, 85);
            styleTitle.Font.Bold = true;

            // custom style data bold
            Style styleDataBold = doc.Styles.Add(StyleType.Paragraph, "DataBold");
            styleDataBold.Font.Size = 14;
            styleDataBold.Font.Name = "Open Sans";
            styleDataBold.Font.Color = Color.FromArgb(34, 34, 34);
            styleDataBold.Font.Bold = true;

            // style name data
            Style styleNameData = doc.Styles.Add(StyleType.Paragraph, "Data");
            styleNameData.Font.Size = 14;
            styleNameData.Font.Color = Color.FromArgb(34, 34, 34);
            styleNameData.Font.Name = "Open Sans";

            //Textbox 1 == header
            Shape textBoxShapeHeader = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 72,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShapeHeader.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeHeader.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeHeader.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph pHeader = new Paragraph(doc);
            pHeader.Runs.Add(new Run(doc, header));
            pHeader.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            pHeader.ParagraphFormat.Style = doc.Styles["Header"];
            pHeader.ParagraphFormat.SpaceBefore = 18;
            textBoxShapeHeader.AppendChild(pHeader);

            builder.InsertNode(textBoxShapeHeader);

            DocumentStyleService.Name(doc, builder, "YGO Name", "data fake", 24, 129);
            DocumentStyleService.Name(doc, builder, "Description", "data fake descriptions", 512, 129);

            Shape textBoxShapeEscort = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 90,
                Left = marginLeft,
                Top = 197
            };
            textBoxShapeEscort.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeEscort.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeEscort.FillColor = Color.FromArgb(255, 255, 255);
            textBoxShapeEscort.Stroked = false;

            Paragraph pEscort = new Paragraph(doc);
            pEscort.Runs.Add(new Run(doc, escort));
            pEscort.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pEscort.ParagraphFormat.SpaceBefore = 3;
            pEscort.ParagraphFormat.Style = doc.Styles["Title"];

            textBoxShapeEscort.AppendChild(pEscort);

            Paragraph pEscortData = new Paragraph(doc);
            pEscortData.Runs.Add(new Run(doc, escortData));
            pEscortData.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pEscortData.ParagraphFormat.SpaceBefore = 8;
            pEscortData.ParagraphFormat.Style = doc.Styles["Data"];

            textBoxShapeEscort.AppendChild(pEscortData);

            Paragraph pEscortSign = new Paragraph(doc);
            pEscortSign.Runs.Add(new Run(doc, escort));
            pEscortSign.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pEscortSign.ParagraphFormat.SpaceBefore = 14;
            pEscortSign.ParagraphFormat.Style = doc.Styles["Title"];

            textBoxShapeEscort.AppendChild(pEscortSign);

            builder.InsertNode(textBoxShapeEscort);

            // insert sign
            Shape textboxSign = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 134,
                Left = marginLeft,
                Top = 294
            };
            textboxSign.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textboxSign.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textboxSign.FillColor = Color.FromArgb(255, 255, 255);
            textboxSign.Fill.Opacity = 0;
            Stroke st = textboxSign.Stroke;
            st.DashStyle = DashStyle.Dash;
            st.Color = Color.FromArgb(183, 183, 183);
            st.Weight = 1;

            // insert image sign
            //var streamImg = new MemoryStream();
            builder.InsertImage("C:\\aspose\\img\\imgs4\\signHandover.png",
            RelativeHorizontalPosition.Page,
            470,
            RelativeVerticalPosition.Page,
            342,
            100,
            60,
            WrapType.Square);

            builder.InsertNode(textboxSign);

            DocumentStyleService.Name(doc, builder, "Escort Person Name", "data fake Escort Person Name", 24, 444);
            DocumentStyleService.Name(doc, builder, "Date Time of Signature", "data fake Date Time of Signature", 512, 444);
            DocumentStyleService.Name(doc, builder, "Restraints Applied", "data fake Escort Person Name", 24, 520);
            DocumentStyleService.Name(doc, builder, "Special Medical Instruction", "data fake Escort Person Name", 512, 520);
            DocumentStyleService.Name(doc, builder, "Accompanied by parents/legal guardians", "data fake Escort Person Name", 24, 576);



            MemoryStream ms = new MemoryStream();
            doc.Save("C:\\aspose\\document\\tempHandover.docx");
            builder.Document.Save(ms, SaveFormat.Pdf);
            byte[] result = ms.ToArray();
            //return result;
            return Ok();
        }

        [HttpGet()]
        [Route("restricted")]
        public async Task<ActionResult> Restricted()
        {
            double marginLeft = 24;
            double marginRight = 24;
            double topStatic = 24;
            string header = "Restricted";
            string titleProfile = "Profile Report On Danielle Lee Tien Hui";
            string titleIntel = "Intel Comments";
            string intelData = "Lorem ipsum dolor sit amet, consetetur sadipscing elitr," +
                " sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat," +
                " sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum." +
                " Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet." +
                " Lorem ipsum dolor sit amet, consetetur sadipscing elitr," +
                " sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat," +
                " sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum." +
                " Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet." +
                " Lorem ipsum dolor sit amet, consetetur sadipscing elitr," +
                " sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat," +
                " sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum." +
                " Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet." +
                " Lorem ipsum dolor sit amet, consetetur sadipscing elitr," +
                " sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat," +
                " sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum." +
                " Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet." +
                " Lorem ipsum dolor sit amet, consetetur sadipscing elitr," +
                " sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat," +
                " sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem";

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            double width = 1000;
            double height = 872;
            builder.PageSetup.PageWidth = width;
            builder.PageSetup.PageHeight = height;

            // Create style header
            Style style = doc.Styles.Add(StyleType.Paragraph, "Header");
            style.Font.Size = 18;
            style.Font.Name = "Open Sans";
            style.Font.Bold = true;
            style.Font.Color = Color.FromArgb(255, 255, 255);

            // custom style title... color down and bold
            Style styleTitle = doc.Styles.Add(StyleType.Paragraph, "Label");
            styleTitle.Font.Size = 14;
            styleTitle.Font.Name = "Open Sans";
            styleTitle.Font.Color = Color.FromArgb(85, 85, 85);
            styleTitle.Font.Bold = true;

            // custom style title
            Style styleDataBold = doc.Styles.Add(StyleType.Paragraph, "Title");
            styleDataBold.Font.Size = 14;
            styleDataBold.Font.Name = "Open Sans";
            styleDataBold.Font.Color = Color.White;

            // style name data
            Style styleNameData = doc.Styles.Add(StyleType.Paragraph, "Data");
            styleNameData.Font.Size = 14;
            styleNameData.Font.Color = Color.FromArgb(34, 34, 34);
            styleNameData.Font.Name = "Open Sans";

            //Textbox 1 == header
            Shape textBoxShapeHeader = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 72,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShapeHeader.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeHeader.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeHeader.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph pHeader = new Paragraph(doc);
            pHeader.Runs.Add(new Run(doc, header));
            pHeader.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            pHeader.ParagraphFormat.Style = doc.Styles["Header"];
            pHeader.ParagraphFormat.SpaceBefore = 16;
            textBoxShapeHeader.AppendChild(pHeader);

            builder.InsertNode(textBoxShapeHeader);

            //insert title
            IntelServices.HeaderShape(doc, builder, titleProfile, 120, 288);
            IntelServices.HeaderShape(doc, builder, titleIntel, 374, 288);

            //insert image profile
            builder.InsertImage("C:\\aspose\\img\\imgs4\\avataGirl.png",
                RelativeHorizontalPosition.Page,
                marginLeft,
                RelativeVerticalPosition.Page,
                120,
                240,
                240,
                WrapType.Square);

            //insert name
            IntelServices.ShapeName(doc, builder, "Name", "Lee tien Hui", 187, 288, 52);
            IntelServices.ShapeName(doc, builder, "NRIC / FIN / Passport Number", "M2356082E", 187, 644, 52);
            IntelServices.ShapeName(doc, builder, "Age", "16 Years 5 Months", 243, 288, 33);
            IntelServices.ShapeName(doc, builder, "DOB", "20 Jun 2007", 243, 644, 33);
            IntelServices.ShapeName(doc, builder, "Case Type", "Remand FGO", 280, 288, 33);
            IntelServices.ShapeName(doc, builder, "Gang Afillation", "Angry Cot Gong", 280, 644, 33);
            IntelServices.ShapeName(doc, builder, "DOA", "20 Jun 2022", 317, 288, 33);
            IntelServices.ShapeName(doc, builder, "UDOD", "20 Jun 2022", 317, 644, 33);

            Shape intelShape = new Shape(doc, ShapeType.TextBox)
            {
                Width = 688,
                Height = 358,
                Left = 288,
                Top = 441
            };
            intelShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            intelShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            intelShape.FillColor = Color.FromArgb(255, 255, 255);
            intelShape.Stroked = false;

            Paragraph pIntel = new Paragraph(doc);
            pIntel.Runs.Add(new Run(doc, intelData));
            pIntel.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pIntel.ParagraphFormat.SpaceBefore = 3;
            pIntel.ParagraphFormat.Style = doc.Styles["Data"];

            intelShape.AppendChild(pIntel);
            builder.InsertNode(intelShape);

            // insert fake footer
            IntelServices.ShapeName(doc, builder, "Intel Officer Name", "Michele Chan", 815, 288, 33);
            IntelServices.ShapeName(doc, builder, "Date Time", "20 Jun 2023 11:30", 815, 644, 33);


            
            MemoryStream ms = new MemoryStream();
            doc.Save("C:\\aspose\\document\\tempIntel.docx");
            builder.Document.Save(ms, SaveFormat.Doc);
            //byte[] result = ms.ToArray();
            return Ok();
        }

        [HttpGet()]
        [Route("intel")]
        public async Task<ActionResult> Intel()
        {
            return Ok();
        }

        [HttpGet()]
        [Route("downloadPdf")]
        public async Task<ActionResult> DownloadPdf(int? id)
        {
            string docDir = "C:\\aspose\\download\\";
            string top1 = "this is top 11111111";
            string top2 = "this is top 2222222222222222222222";
            string top3 = "this is top 3";
            string a = $"{top1}\n{top2}\n{top3}";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            var width = doc.GetPageInfo(0).SizeInPoints.Width;
            var height = doc.GetPageInfo(0).SizeInPoints.Height;
            // Insert another textbox with specific margins.
            //Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, RelativeHorizontalPosition.Page, 50, RelativeVerticalPosition.Page, 50, width - 100, 100, WrapType.None);
            Shape textBoxShape =new Shape(doc, ShapeType.TextBox)
            {
                Width = width -100,
                Height = 200,
                Left = 50,
                Top = 50
            };
            textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            TextBox textBox = textBoxShape.TextBox;
            textBoxShape.FillColor = Color.Gray;
            //textBox.InternalMarginLeft = 50;
            textBox.VerticalAnchor = TextBoxAnchor.Middle;

            
            for(int i= 0; i < 3; i++)
            {
                Paragraph p = new Paragraph(doc);
                p.Runs.Add(new Run(doc, i + top1));
                textBoxShape.AppendChild(p);
                p.ParagraphFormat.LineSpacing = 20;
                p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                p.ParagraphFormat.Style.Font.Color = Color.Red;
                p.ParagraphFormat.Style.Font.Size = 40;
                p.ParagraphFormat.Style.Font.Bold = true;
                p.ParagraphFormat.Style.Font.Name = "Calibri";
                //p.ParagraphBreakFont.Size = 40;
                //p.ParagraphBreakFont.Bold = true;
                //p.ParagraphBreakFont.Name = "Calibri";
                if(i == 2)
                {
                    //p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    //p.ParagraphFormat.Style.Font.Color = Color.White;
                    //p.ParagraphFormat.Style.Font.Size = 20;
                    //p.ParagraphFormat.Style.Font.Bold = false;
                    //p.ParagraphFormat.Style.Font.Name = "Cambria";
                    //p.ParagraphBreakFont.Color = Color.White;
                    //p.ParagraphBreakFont.Size = 20;
                    //p.ParagraphBreakFont.Bold = false;
                    //p.ParagraphBreakFont.Name = "Cambria";
                }
            }
            //builder.InsertNode(textBoxShape);

            //Paragraph p1 = new Paragraph(doc);
            //p1.Runs.Add(new Run(doc, top2));
            //p1.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            //p1.ParagraphFormat.Style.Font.Color = Color.White;
            //p1.ParagraphFormat.Style.Font.Size = 20;
            //p1.ParagraphFormat.Style.Font.Bold = false;
            //p1.ParagraphFormat.Style.Font.Name = "Cambria";
            //textBoxShape.AppendChild(p1);
            builder.InsertNode(textBoxShape);

            Shape textBoxShape1 = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - 100,
                Height = 200,
                Left = 50,
                Top = 300
            };
            textBoxShape1.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShape1.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            TextBox textBox1 = textBoxShape1.TextBox;
            textBoxShape1.FillColor = Color.Gray;

            Paragraph p1 = new Paragraph(doc);
            p1.Runs.Add(new Run(doc, top2));
            p1.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            p1.ParagraphFormat.Style.Font.Color = Color.White;
            p1.ParagraphFormat.Style.Font.Size = 20;
            p1.ParagraphFormat.Style.Font.Bold = false;
            p1.ParagraphFormat.Style.Font.Name = "Cambria";
            textBoxShape1.AppendChild(p1);

            builder.InsertNode(textBoxShape1);

            builder.MoveToSection(0);
            //builder.InsertNode(textBoxShape);
            //builder.MoveToSection(0);
            //builder.MoveTo(textBoxShape.LastParagraph);

            doc.Save(docDir + "check.docx");
            Document docPdf = new Document(docDir + "check.docx");
            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.Pdf17;
            docPdf.Save(docDir + "checkPdf.pdf", options);
            var dochtml = new Document(docDir + "html template\\inline.html");
            dochtml.Save(docDir + "Output.docx");
            return Ok();
        }


        [HttpGet()]
        [Route("downloadPdf222")]
        public async Task<ActionResult> DownloadPdf222(int? id)
        {
            Double marginLeft = 24;
            Double marginRight = 24;
            Double between = 15;
            //position top for shape
            Double topStatic = 24;
            //position shape left
            Double marginLeftStatic = 24;

            string docImg = "C:\\aspose\\download\\signature example\\";
            string docDir = "C:\\aspose\\download\\";
            string headerParent = "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT";
            string headerChild = "RECORD OF DEPOSIT/ WITHDRAWAL OF PERSONAL EFFECTS OF RESIDENT";
            string headerDate = "P123A2023080001";
            string titleFirst = "Particulars of Resident";
            string titleSecond = "To be completed when resident’s personal effects are deposited";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.PageWidth = 800;
            builder.PageSetup.PageHeight = 975;
            var width = doc.GetPageInfo(0).SizeInPoints.Width;
            var height = doc.GetPageInfo(0).SizeInPoints.Height;

            // Create style header parent
            Style style = doc.Styles.Add(StyleType.Paragraph, "HeaderParent");
            style.Font.Size = 18;
            style.Font.Name = "Open Sans";
            style.Font.Bold = true;
            //style.Font.Spacing = 2;
            style.Font.Color = Color.FromArgb(255, 255, 255);
            //style.Font.Spacing = 20;

            // Create style header child
            Style styleChild = doc.Styles.Add(StyleType.Paragraph, "HeaderChild");
            styleChild.Font.Size = 14;
            styleChild.Font.Name = "Open Sans";
            styleChild.Font.Bold = false;
            //style.Font.Spacing = 1;
            styleChild.Font.Color = Color.FromArgb(255, 255, 255);

            // Create style header date
            Style styleDate = doc.Styles.Add(StyleType.Paragraph, "HeaderDate");
            styleDate.Font.Size = 14;
            styleDate.Font.Name = "Open Sans";
            styleDate.Font.Bold = false;
            styleDate.Font.Color = Color.FromArgb(255, 255, 255);

            // custom style title
            Style styleHeader = doc.Styles.Add(StyleType.Paragraph, "Title");
            styleHeader.Font.Size = 16;
            styleHeader.Font.Name = "Open Sans";
            styleHeader.Font.Color = Color.FromArgb(255, 255, 255);
            styleHeader.Font.Bold = true;

            // style name 
            Style styleName = doc.Styles.Add(StyleType.Paragraph, "StyleName");
            styleName.Font.Size = 14;
            styleName.Font.Color = Color.FromArgb(14, 37, 84);
            styleName.Font.Name = "Open Sans";
            styleName.Font.Bold = true;


            // style name 
            Style styleNameData = doc.Styles.Add(StyleType.Paragraph, "StyleNameData");
            styleNameData.Font.Size = 14;
            styleNameData.Font.Color = Color.FromArgb(14, 37, 84);
            styleNameData.Font.Name = "Open Sans";

            // Create a list and make sure the paragraphs that use this style will use this list.
            //style.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
            //style.ListFormat.ListLevelNumber = 0;


            //Textbox 1 == header
            Shape textBoxShape = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 110,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShape.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph p = new Paragraph(doc);
            p.Runs.Add(new Run(doc, headerParent));
            p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            p.ParagraphFormat.Style = doc.Styles["HeaderParent"];
            p.ParagraphFormat.SpaceBefore = 24;
            textBoxShape.AppendChild(p);

            Paragraph p2 = new Paragraph(doc);
            p2.Runs.Add(new Run(doc, headerChild));
            p2.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            p2.ParagraphFormat.SpaceBefore = 8;
            p2.ParagraphFormat.Style = doc.Styles["HeaderChild"];
            textBoxShape.AppendChild(p2);

            Paragraph p1 = new Paragraph(doc);
            p1.Runs.Add(new Run(doc, headerDate));
            p1.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            p1.ParagraphFormat.SpaceBefore = 8;
            p1.ParagraphFormat.Style = doc.Styles["HeaderDate"];
            textBoxShape.AppendChild(p1);

            builder.InsertNode(textBoxShape);
            topStatic += textBoxShape.Height + 24;



            // add title first
            Shape textBoxShapeTittle = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 38,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShapeTittle.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeTittle.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeTittle.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph pTitle = new Paragraph(doc);
            pTitle.Runs.Add(new Run(doc, titleFirst));
            pTitle.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pTitle.ParagraphFormat.SpaceBefore = 8;
            pTitle.ParagraphFormat.Style = doc.Styles["Title"];
            textBoxShapeTittle.AppendChild(pTitle);

            builder.InsertNode(textBoxShapeTittle);
            topStatic += textBoxShapeTittle.Height + 16;

            // Apply the paragraph style to the document builder's current paragraph, and then add some text.
            //builder.ParagraphFormat.Style = style;
            //builder.Writeln("Hello World: MyStyle1, bulleted list.");




            // Change the document builder's style to one that has no list formatting and write another paragraph.
            //builder.ParagraphFormat.Style = doc.Styles["Normal"];
            //builder.Writeln("Hello World: Normal.");


            //builder.InsertNode(textBoxShape);


            //Shape textBoxShape1 = new Shape(doc, ShapeType.TextBox)
            //{
            //    Width = width - 100,
            //    Height = 200,
            //    Left = 50,
            //    Top = 300
            //};
            //textBoxShape1.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            //textBoxShape1.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            //TextBox textBox1 = textBoxShape1.TextBox;
            //textBoxShape1.FillColor = Color.Gray;




            // create 4 cell name
            Shape textBoxShapeName1 = new Shape(doc, ShapeType.TextBox)
            {
                Width = 149,
                Height = 33,
                Left = marginLeftStatic,
                Top = topStatic
            };
            textBoxShapeName1.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeName1.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeName1.Stroked = false;
            textBoxShapeName1.FillColor = Color.FromArgb(234, 239, 244);

            Paragraph pName1 = new Paragraph(doc);
            pName1.Runs.Add(new Run(doc, "Name"));
            pName1.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pName1.ParagraphFormat.Style = doc.Styles["StyleName"];
            textBoxShapeName1.AppendChild(pName1);
            builder.InsertNode(textBoxShapeName1);
            marginLeftStatic += textBoxShapeName1.Width + 8;

            // add 3 cell other
            Shape textBoxShapeName2 = new Shape(doc, ShapeType.TextBox)
            {
                Width = 199,
                Height = 33,
                Left = marginLeftStatic,
                Top = topStatic
            };
            textBoxShapeName2.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeName2.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeName2.Stroked = false;
            textBoxShapeName2.FillColor = Color.White;

            Paragraph pName2 = new Paragraph(doc);
            pName2.Runs.Add(new Run(doc, "Lucas with lucian"));
            pName2.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pName2.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            textBoxShapeName2.AppendChild(pName2);
            builder.InsertNode(textBoxShapeName2);
            marginLeftStatic += textBoxShapeName2.Width + 72;

            // cell 3

            Shape textBoxShapeName3 = new Shape(doc, ShapeType.TextBox)
            {
                Width = 149,
                Height = 33,
                Left = marginLeftStatic,
                Top = topStatic
            };
            textBoxShapeName3.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeName3.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeName3.Stroked = false;
            textBoxShapeName3.FillColor = Color.FromArgb(234, 239, 244);

            Paragraph pName3 = new Paragraph(doc);
            pName3.Runs.Add(new Run(doc, "NRIC / FIN"));
            pName3.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pName3.ParagraphFormat.Style = doc.Styles["StyleName"];
            textBoxShapeName3.AppendChild(pName3);
            builder.InsertNode(textBoxShapeName3);
            marginLeftStatic += textBoxShapeName3.Width + 8;

            // cell 4

            Shape textBoxShapeName4 = new Shape(doc, ShapeType.TextBox)
            {
                Width = 199,
                Height = 33,
                Left = marginLeftStatic,
                Top = topStatic
            };
            textBoxShapeName4.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeName4.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeName4.Stroked = false;
            textBoxShapeName4.FillColor = Color.White;

            Paragraph pName4 = new Paragraph(doc);
            pName4.Runs.Add(new Run(doc, "S1000HST99900"));
            pName4.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pName4.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            textBoxShapeName4.AppendChild(pName4);
            builder.InsertNode(textBoxShapeName4);

            // end cell reset marginLeftStatic to default
            marginLeftStatic = 24;

            // add top under 4 cell name
            topStatic += textBoxShapeName4.Height + 24;


            // tile Second
            Shape textBoxShapeTittle2 = new Shape(doc, ShapeType.TextBox)
            {
                Width = width - marginLeft - marginRight,
                Height = 38,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShapeTittle2.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeTittle2.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeTittle2.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph pTitle2 = new Paragraph(doc);
            pTitle2.Runs.Add(new Run(doc, titleSecond));
            pTitle2.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pTitle2.ParagraphFormat.SpaceBefore = 8;
            pTitle2.ParagraphFormat.Style = doc.Styles["Title"];
            textBoxShapeTittle2.AppendChild(pTitle2);

            builder.InsertNode(textBoxShapeTittle2);
            topStatic += textBoxShapeTittle2.Height + 16;

            //Shape shapeImg = new Shape(doc, ShapeType.Image)
            //{
            //    Width = 100,
            //    Left = 50,
            //    Top = topStatic,
            //    Height = 50,
            //};
            //shapeImg.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            //shapeImg.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            //Shape imgShape = builder.InsertImage(docImg + "sign 1.png", RelativeHorizontalPosition.Page,
            //    100,
            //    RelativeVerticalPosition.Page,
            //    500,
            //    150,
            //    100,
            //    WrapType.None);

            //imgShape.Font.Border.Color = Color.Green;
            //imgShape.Font.Border.LineWidth = 5;
            //imgShape.Font.Border.LineStyle = LineStyle.Dot;

            //Stroke strokeImg = imgShape.Stroke;
            //strokeImg.On = true;
            //strokeImg.EndCap = EndCap.Square;
            //strokeImg.Weight = 2;
            //strokeImg.LineStyle = ShapeLineStyle.ThinThick;
            //strokeImg.Color = Color.Blue;

            //DocumentStyleService.HeaderShape(doc,ref builder, "", 40);
            //DocumentStyleService.HeaderShape(doc,ref builder, "", 40+218);
            //DocumentStyleService.HeaderShape(doc,ref builder, "", 40+218+218);
            // insert shape sign 
            //Shape textBoxShapeSign = new Shape(doc, ShapeType.TextBox)
            //{
            //    Width = 218,
            //    Height = 130,
            //    Left = 40,
            //    Top = topStatic
            //};
            //textBoxShapeSign.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            //textBoxShapeSign.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            ////textBoxShapeSign.Stroked = false;
            //textBoxShapeSign.FillColor = Color.White;

            //Paragraph pSign1 = new Paragraph(doc);
            //pSign1.Runs.Add(new Run(doc, "I confirm that all the items as stated in the following pages are correct."));
            //pSign1.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pSign1.ParagraphFormat.Style = doc.Styles["StyleName"];
            //pSign1.ParagraphFormat.SpaceBefore = 8;
            //textBoxShapeSign.AppendChild(pSign1);

            //Paragraph pSign2 = new Paragraph(doc);
            //pSign2.Runs.Add(new Run(doc, "Lucas with lucian"));
            //pSign2.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pSign2.ParagraphFormat.SpaceBefore = 16;
            //pSign2.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            //textBoxShapeSign.AppendChild(pSign2);

            //Paragraph pSign3 = new Paragraph(doc);
            //pSign3.Runs.Add(new Run(doc, "Date: 6/19/2008 7:00:00 AM"));
            //pSign3.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //pSign3.ParagraphFormat.SpaceBefore = 8;
            //pSign3.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            //textBoxShapeSign.AppendChild(pSign3);

            //builder.InsertNode(textBoxShapeSign);
            //topStatic += textBoxShapeSign.Height + 8;

            // insert img

            builder.InsertImage(docImg + "sign 1.png",
            RelativeHorizontalPosition.Page,
            marginLeft,
            RelativeVerticalPosition.Page,
            topStatic,
            (width - marginLeft - marginRight - 2 * between) / 3,
            60,
            WrapType.Square);

            // test insert image border dot
            //builder.MoveToDocumentEnd();
            //builder.MoveTo(textBoxShapeName4);
            //builder.Writeln("this is cursor move to test");
            //Shape shape = new Shape(doc, ShapeType.Image);
            ////{
            ////    Width = 200,
            ////    Height = 50,
            ////};
            //shape.ImageData.ImageBytes = System.IO.File.ReadAllBytes(docImg + "sign 1.png");
            //shape.WrapType = WrapType.Inline;
            ////shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            ////shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            //shape.ImageData.Borders.Color = Color.Red;
            //shape.ImageData.Borders.LineStyle = LineStyle.Dot;
            //shape.ImageData.Borders.LineWidth = 2;

            //textBoxShapeSign.AppendChild(shape);
            //builder.InsertNode(shape);
            // Set borders for a shape.
            //doc.Save(docDir + "check11112.docx");
            // test textBox 2

            //for(int i= 0; i < 10; i++)
            //{
            //Shape textBoxShape1 = new Shape(doc, ShapeType.TextBox)
            //{
            //    Width = width - 100,
            //    Height = 200,
            //    Left = 50,
            //    Top = topStatic
            //};
            //textBoxShape1.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            //textBoxShape1.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            //TextBox textBox1 = textBoxShape1.TextBox;
            //textBoxShape1.FillColor = Color.Blue;

            //Paragraph p3 = new Paragraph(doc);
            //p3.Runs.Add(new Run(doc, top2));
            //p3.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //p3.ParagraphFormat.Style = style;

            //textBoxShape1.AppendChild(p3);

            //builder.InsertNode(textBoxShape1);
            //    topStatic += textBoxShape1.Height;
            //}

            Document dstDoc = new Document();
            dstDoc = doc.Clone();
            //dstDoc.FirstSection.Body.AppendParagraph("Destination document text. ");

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Text goes here...");
            //add footer having current date
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.InsertField("Date", "");

            // Append the source document to the destination document while preserving its formatting,
            // then save the source document to the local file system.
            doc.AppendDocument(dstDoc, ImportFormatMode.KeepSourceFormatting);

            builder.MoveToSection(1);
            topStatic = 50;


            //Shape textBoxShape2 = new Shape(doc, ShapeType.TextBox)
            //{
            //    Width = width - marginLeft - marginRight,
            //    Height = 100,
            //    Left = marginLeft,
            //    Top = topStatic,
            //};
            //textBoxShape2.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            //textBoxShape2.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            //textBoxShape2.FillColor = Color.FromArgb(14, 37, 84);

            //Paragraph p11 = new Paragraph(doc);
            //p11.Runs.Add(new Run(doc, headerParent));
            //p11.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            //p11.ParagraphFormat.Style = doc.Styles["HeaderParent"];
            //p11.ParagraphFormat.SpaceBefore = 15;
            //textBoxShape2.AppendChild(p11);

            //Paragraph p21 = new Paragraph(doc);
            //p21.Runs.Add(new Run(doc, headerChild));
            //p21.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            //p21.ParagraphFormat.SpaceBefore = 8;
            //p21.ParagraphFormat.Style = doc.Styles["HeaderChild"];
            //textBoxShape2.AppendChild(p21);

            //Paragraph p12 = new Paragraph(doc);
            //p12.Runs.Add(new Run(doc, headerDate));
            //p12.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            //p12.ParagraphFormat.SpaceBefore = 8;
            //p12.ParagraphFormat.Style = doc.Styles["HeaderDate"];
            //textBoxShape2.AppendChild(p12);

            //builder.InsertNode(textBoxShape);
            topStatic += textBoxShape.Height + 30;


            //builder.InsertField("Date", "");
            var currentSection = builder.CurrentSection;
            Section previousSection = (Section)currentSection.PreviousSibling;


            //currentSection.HeadersFooters.Clear();

            //foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
            //    currentSection.HeadersFooters.Add(headerFooter.Clone(true));

            builder.Document.Save(docDir + "property.docx");
            Document docPdf = new Document(docDir + "property.docx");
            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.PdfA4;
            docPdf.Save(docDir + "property.pdf", options);
            return Ok();
        }

        [HttpPost()]
        [Route("testNImage")]
        public async Task<ActionResult> TestImage(int? a)
        {
            Double marginLeft = 24;
            Double marginRight = 24;
            Double topStatic = 24;
            Double marginLeftStatic = 24;

            string docSign = "C:\\aspose\\download\\signature example\\";
            string docImg = "C:\\aspose\\download\\item\\";
            string docDir = "C:\\aspose\\download\\";
            string headerParent = "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT" +
                "MINISTRY OF SOCIAL AND FAMILY DEVELOPMENT";
            string headerChild = "RECORD OF DEPOSIT/ WITHDRAWAL OF PERSONAL EFFECTS OF RESIDENT";
            string headerDate = "P123A2023080001";
            string titleFirst = "Particulars of Resident";
            string titleSecond = "To be completed when resident’s personal effects are deposited";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.PageWidth = 800;
            builder.PageSetup.PageHeight = 975;
            var width = doc.GetPageInfo(0).SizeInPoints.Width;
            var height = doc.GetPageInfo(0).SizeInPoints.Height;

            // Create style header parent
            Style style = doc.Styles.Add(StyleType.Paragraph, "HeaderParent");
            style.Font.Size = 18;
            style.Font.Name = "Open Sans";
            style.Font.Bold = true;
            //style.Font.Spacing = 2;
            style.Font.Color = Color.FromArgb(255, 255, 255);
            //style.Font.Spacing = 20;

            // Create style header child
            Style styleChild = doc.Styles.Add(StyleType.Paragraph, "HeaderChild");
            styleChild.Font.Size = 14;
            styleChild.Font.Name = "Open Sans";
            styleChild.Font.Bold = false;
            //style.Font.Spacing = 1;
            styleChild.Font.Color = Color.FromArgb(255, 255, 255);

            // Create style header date
            Style styleDate = doc.Styles.Add(StyleType.Paragraph, "HeaderDate");
            styleDate.Font.Size = 14;
            styleDate.Font.Name = "Open Sans";
            styleDate.Font.Bold = false;
            styleDate.Font.Color = Color.FromArgb(255, 255, 255);

            // custom style title
            Style styleHeader = doc.Styles.Add(StyleType.Paragraph, "Title");
            styleHeader.Font.Size = 16;
            styleHeader.Font.Name = "Open Sans";
            styleHeader.Font.Color = Color.FromArgb(255, 255, 255);
            styleHeader.Font.Bold = true;

            // style name 
            Style styleName = doc.Styles.Add(StyleType.Paragraph, "StyleName");
            styleName.Font.Size = 14;
            styleName.Font.Color = Color.FromArgb(14, 37, 84);
            styleName.Font.Name = "Open Sans";
            styleName.Font.Bold = true;


            // style name 
            Style styleNameData = doc.Styles.Add(StyleType.Paragraph, "StyleNameData");
            styleNameData.Font.Size = 14;
            styleNameData.Font.Color = Color.FromArgb(14, 37, 84);
            styleNameData.Font.Name = "Open Sans";
            //Textbox 1 == header
            Shape textBoxShape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = width - marginLeft - marginRight,
                Left = marginLeft,
                Top = topStatic
            };
            textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShape.FillColor = Color.FromArgb(14, 37, 84);
            textBoxShape.TextBox.FitShapeToText = true;

            Paragraph p = new Paragraph(doc);
            p.Runs.Add(new Run(doc, headerParent));
            p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            p.ParagraphFormat.Style = doc.Styles["HeaderParent"];
            p.ParagraphFormat.SpaceBefore = 24;
            var widd = p.FrameFormat.Width;
            var heiii = p.FrameFormat.Height;
            var tem = p.FrameFormat.RelativeVerticalPosition;

            textBoxShape.AppendChild(p);

            builder.InsertNode(textBoxShape);
            topStatic += textBoxShape.Height + 24;

            Node elementNode = textBoxShape.Clone(true);
            builder.InsertNode(elementNode);


            DocumentStyleService.HeaderShape(doc,  builder, "", 40);
            DocumentStyleService.HeaderShape(doc,  builder, "", 40 + 218);
            DocumentStyleService.HeaderShape(doc,  builder, "", 40 + 218 + 218);

            // test line para
            //var collector = new LayoutCollector(doc);
            //var it = new LayoutEnumerator(doc);
            //var listNode = doc.GetChildNodes(NodeType.Paragraph, true);
            //foreach (Paragraph paragraph in listNode)
            //{
            //    var paraBreak = collector.GetEntity(paragraph);

            //    object stop = null;
            //    var prevItem = paragraph.PreviousSibling;
            //    if (prevItem != null)
            //    {
            //        var prevBreak = collector.GetEntity(prevItem);
            //        if (prevItem is Paragraph)
            //        {
            //            it.Current = collector.GetEntity(prevItem); // para break
            //            it.MoveParent();    // last line
            //            stop = it.Current;
            //        }
            //        else
            //        {
            //            throw new Exception();
            //        }
            //    }
            //    var ss = paragraph.Count;
            //    it.Current = paraBreak;
            //    it.MoveParent();

            //    // We move from line to line in a paragraph.
            //    // When paragraph spans multiple pages the we will follow across them.
            //    var count = 1;
            //    while (it.Current != stop)
            //    {
            //        if (!it.MovePreviousLogical())
            //            break;
            //        count++;
            //    }

            //    var paraText = paragraph.GetText();

            //    Console.WriteLine($"Paragraph '{paraText}' has {ss} line(-s).");



            //}

            //RenderedDocument layoutDoc = new RenderedDocument(doc);
            //foreach (Section sections in doc.Sections)
            //{
            //    foreach (Paragraph paragraph in sections.Body.Paragraphs)
            //    {
            //        Console.WriteLine("Paragraph text : " + paragraph.ToString(SaveFormat.Text));
            //        Console.WriteLine("Paragraph lines : " + layoutDoc.GetLayoutEntitiesOfNode(paragraph).Count);
            //        Console.WriteLine("Words Count : " + paragraph.ToString(SaveFormat.Text).Split(' ').Length);
            //    }
            //}

            builder.Document.Save(docDir + "property.docx");
            Document docPdf = new Document(docDir + "property.docx");
            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.Pdf17;
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;
            DocumentBuilder builderPdf = new DocumentBuilder(docPdf);
            Aspose.Words.Font f = builderPdf.Font;
            builderPdf.Document.Save(docDir + "property.pdf", options);
            return Ok();
        }

        [HttpPost]
        [Route("downloadformatpdf")]
        public async Task<ActionResult> TestFormatPdf(int? Id)
        {
            //string docDir = @"";
            string docDir = "C:\\aspose\\download\\";
            Document docPdf = new Document(docDir + "property.docx");
            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.Pdf17;
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;
            options.UseCoreFonts = true;
            DocumentBuilder builderPdf = new DocumentBuilder(docPdf);
            Aspose.Words.Font f = builderPdf.Font;
            builderPdf.Document.Save(docDir + "property.pdf", options);
            //FontSettings FontSettings = new FontSettings();
            //FontSettings.SetFontSubstitutes("MS Gothic", new string[] { "Arial Unicode MS" });
            //docPdf.FontSettings =  

            docPdf.Save(docDir + "property.pdf", SaveFormat.Pdf);
            return Ok();
        }
    }
}
