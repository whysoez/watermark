using Microsoft.AspNetCore.Mvc;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;
using Aspose.Words.Drawing;
using System.Text;

namespace watermark.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GSRController : Controller
    {
        [HttpPost()]
        [Route("GSRTable")]
        public async Task<IActionResult> CreateTable()
        {
            //MemoryStream ms = new MemoryStream();
            //await file.CopyToAsync(ms);
            //ms.Seek(0, SeekOrigin.Begin);

            string licenseString = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>\r\n<License>\r\n\t<Data>\r\n\t\t<LicensedTo>AvePoint</LicensedTo>\r\n\t\t<EmailTo>it_billing@avepoint.com</EmailTo>\r\n\t\t<LicenseType>Developer OEM</LicenseType>\r\n\t\t<LicenseNote>1 Developer And Unlimited Deployment Locations</LicenseNote>\r\n\t\t<OrderID>230601004913</OrderID>\r\n\t\t<UserID>336519</UserID>\r\n\t\t<OEM>This is a redistributable license</OEM>\r\n\t\t<Products>\r\n\t\t\t<Product>Aspose.Total for .NET</Product>\r\n\t\t</Products>\r\n\t\t<EditionType>Professional</EditionType>\r\n\t\t<SerialNumber>d1e7b90d-f81e-4ac8-819b-25464e23b292</SerialNumber>\r\n\t\t<SubscriptionExpiry>20240608</SubscriptionExpiry>\r\n\t\t<LicenseVersion>3.0</LicenseVersion>\r\n\t\t<LicenseInstructions>https://purchase.aspose.com/policies/use-license</LicenseInstructions>\r\n\t</Data>\r\n\t<Signature>V8BKQvsO7XLWS2WjYWc0ihX2fyIjkY6/6GU4xB1QvQ6tekEtJHY4GrraaywmKAWJT55qX0U6esMCD7FNujoXHUAAe+Jtx6EFAK+tzMlmqOtt+4VYTa8hBMSlRqZldl/Z9BzoAYGLGaNOZm4MjK6l1yF+TLsa3QbomWzcGmnnimo=</Signature>\r\n</License>";
            using MemoryStream msLic = new MemoryStream(Encoding.UTF8.GetBytes(licenseString));
            {
                Aspose.Words.License license = new Aspose.Words.License();
                license.SetLicense(msLic);
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We start by creating the table object. Note that we must pass the document object
            // to the constructor of each node. This is because every node we create must belong
            // to some document.
            Table table = builder.StartTable();
            //table.SetBorder(BorderType.Left, LineStyle.Double, 2, Color.Red, true);
            //table.SetBorder(BorderType.Top, LineStyle.Double, 2, Color.Red, true);
            //table.SetBorder(BorderType.Right, LineStyle.Double, 2, Color.Red, true);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 24;
            builder.Font.Color = Color.FromArgb(34, 34, 34);
            builder.Font.Name = "Open Sans";
            builder.Font.Bold = true;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            builder.Write("Cell at 40 points width");


            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell at 20% width");

            // Insert a auto sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(60);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.");
            
            builder.EndRow();

            builder.InsertCell();

            RowFormat rowformat = builder.RowFormat;
            rowformat.Height = 100;
            rowformat.HeightRule = HeightRule.Exactly;
            builder.Font.Size = 20;
            builder.Font.Color = Color.FromArgb(192, 192, 192);
            builder.Font.Name = "Times New Roman";
            builder.Font.Bold = false;

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            builder.Write("Cell at 40 points width");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Write("Cell at 40 points width");

            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell at 20% width");

            // Insert a auto sized cell.
            //builder.InsertCell();
            //builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            //builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            //builder.Write("In this case the cell will fill up the rest of the available space.");
            builder.EndRow();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.EndRow();
            table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table.SetBorder(BorderType.Left, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            //table.SetBorder(BorderType.Bottom, LineStyle.None, 1, Color.FromArgb(192, 192, 192), true);
            table.Alignment = TableAlignment.Center;
            table.VerticalAnchor = RelativeVerticalPosition.Page;
            table.HorizontalAnchor = RelativeHorizontalPosition.Page;
            table.AbsoluteVerticalDistance = 50;
            table.AbsoluteHorizontalDistance = 50;
            builder.EndTable();

            //builder.Write("blank text test");

            //insert table2
            Table table2 = builder.StartTable();
            //table.SetBorder(BorderType.Left, LineStyle.Double, 2, Color.Red, true);
            //table.SetBorder(BorderType.Top, LineStyle.Double, 2, Color.Red, true);
            //table.SetBorder(BorderType.Right, LineStyle.Double, 2, Color.Red, true);
            rowformat.Height = 400;
            rowformat.HeightRule = HeightRule.Exactly;
            var testCell = builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            builder.Write("Cell at 40 points width");

            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell at 20% width");

            // Insert a auto sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.In this case the cell will fill up the rest of the available space.");

            builder.EndRow();
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Write("Cell at 40 points width");

            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell at 20% width");

            // Insert a auto sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("In this case the cell will fill up the rest of the available space.");
            builder.EndRow();
            table2.SetBorder(BorderType.Horizontal, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            //table2.SetBorder(BorderType.Top, LineStyle.None, 1, Color.FromArgb(192, 192, 192), true);
            table2.SetBorder(BorderType.Left, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table2.SetBorder(BorderType.Right, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table2.SetBorder(BorderType.Bottom, LineStyle.Single, 1, Color.FromArgb(192, 192, 192), true);
            table2.Alignment = TableAlignment.Center;
            table2.VerticalAnchor = RelativeVerticalPosition.Paragraph;
            table2.HorizontalAnchor = RelativeHorizontalPosition.Page;
            table2.AbsoluteVerticalDistance = 0;
            table2.AbsoluteHorizontalDistance = 50;
            builder.EndTable();

            Table tableNest = builder.StartTable();
            tableNest.SetShading(TextureIndex.Texture10Percent, Color.AliceBlue, Color.Red);
            rowformat.Height = 100;
            rowformat.HeightRule = HeightRule.Auto;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.InsertCell();
            //builder.Write("insert cell to table");
            Table tableChild = builder.StartTable();
            //rowformat.Height = 200;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            tableChild.SetShading(TextureIndex.Texture15Percent, Color.Red, Color.Red);
            builder.InsertCell();
            builder.Write("insert cell to table test spose.Cells provides a class, Workbook, that represents an Excel file. The Workbook class contains the Worksheets collection that allows access to each worksheet in the Excel file. A worksheet is represented by the Worksheet class.\r\n\r\nThe Worksheet class provides the PageSetup property used to set the page setup options for a worksheet. The PageSetup attribute is an object of the PageSetup class that enables developers to set different page layout options for a printed worksheet. The PageSetup class provides various properties and methods used to set page setup options.\r\n\r\nPage Margins\r\nSet page margins (left, right, top, bottom) using PageSetup class members. A few of the methods are listed below which are used to specify page margins:\r\n\r\nLeftMargin\r\nRightMargin\r\nTopMargin\r\nBottomMarginspose.Cells provides a class, Workbook, that represents an Excel file. The Workbook class contains the Worksheets collection that allows access to each worksheet in the Excel file. A worksheet is represented by the Worksheet class.\r\n\r\nThe Worksheet class provides the PageSetup property used to set the page setup options for a worksheet. The PageSetup attribute is an object of the PageSetup class that enables developers to set different page layout options for a printed worksheet. The PageSetup class provides various properties and methods used to set page setup options.\r\n\r\nPage Margins\r\nSet page margins (left, right, top, bottom) using PageSetup class members. A few of the methods are listed below which are used to specify page margins:\r\n\r\nLeftMargin\r\nRightMargin\r\nTopMargin\r\nBottomMarginspose.Cells provides a class, Workbook, that represents an Excel file. The Workbook class contains the Worksheets collection that allows access to each worksheet in the Excel file. A worksheet is represented by the Worksheet class.\r\n\r\nThe Worksheet class provides the PageSetup property used to set the page setup options for a worksheet. The PageSetup attribute is an object of the PageSetup class that enables developers to set different page layout options for a printed worksheet. The PageSetup class provides various properties and methods used to set page setup options.\r\n\r\nPage Margins\r\nSet page margins (left, right, top, bottom) using PageSetup class members. A few of the methods are listed below which are used to specify page margins:\r\n\r\nLeftMargin\r\nRightMargin\r\nTopMargin\r\nBottomMargin");
            builder.Write("test table child");
            tableChild.SetBorder(BorderType.Top, LineStyle.Single, 2, Color.Red, true);
            tableChild.SetBorder(BorderType.Bottom, LineStyle.Single, 2, Color.Red, true);
            tableChild.VerticalAnchor = RelativeVerticalPosition.Paragraph;
            tableChild.HorizontalAnchor = RelativeHorizontalPosition.Page;
            tableChild.AbsoluteVerticalDistance = 20;
            tableChild.AbsoluteHorizontalDistance = 60;
            builder.EndTable();
            tableNest.VerticalAnchor = RelativeVerticalPosition.Paragraph;
            tableNest.HorizontalAnchor = RelativeHorizontalPosition.Page;
            tableNest.SetBorder(BorderType.Top, LineStyle.Single, 2, Color.Yellow, true);
            tableNest.SetBorder(BorderType.Bottom, LineStyle.Single, 2, Color.Yellow, true);
            tableNest.VerticalAnchor = RelativeVerticalPosition.Paragraph;
            tableNest.HorizontalAnchor = RelativeHorizontalPosition.Page;
            tableNest.AbsoluteVerticalDistance = 50;
            tableNest.AbsoluteHorizontalDistance = 50;
            builder.EndTable();

            // Go to the primary footer
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            // Add fields for current page number
            builder.Write("Page ");
            builder.InsertField("PAGE", "Page");
            // Add any custom text
            builder.Write(" / ");
            // Add field for total page numbers in document
            builder.InsertField("NUMPAGES", "Total Page");

            // Save the document to disk.
            doc.Save("C:\\Downloads\\testGSRResult\\resultTable.doc");
            Document doc2 = new Document("C:\\Downloads\\testGSRResult\\resultTable.doc");
            doc.Save("C:\\Downloads\\testGSRResult\\resultTable.pdf");

            return Ok();
        }
    }
}
