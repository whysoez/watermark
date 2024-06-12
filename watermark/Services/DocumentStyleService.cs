using Aspose.Cells.Charts;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

namespace watermark.Controllers
{
    public static class DocumentStyleService
    {
        public static DocumentBuilder HeaderShape(Document doc,  DocumentBuilder builder, string data, double marginLeft)
        {
            Shape textBoxShapeSign = new Shape(doc, ShapeType.TextBox)
            {
                Width = 150,
                Height = 130,
                Left = marginLeft,
                Top = 500
            };
            textBoxShapeSign.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeSign.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeSign.Stroked = false;
            textBoxShapeSign.FillColor = Color.White;

            Paragraph pSign1 = new Paragraph(doc);
            pSign1.Runs.Add(new Run(doc, "I confirm that all the items as stated in the following pages are correct."));
            pSign1.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pSign1.ParagraphFormat.Style = doc.Styles["StyleName"];
            pSign1.ParagraphFormat.SpaceBefore = 8;
            textBoxShapeSign.AppendChild(pSign1);

            Paragraph pSign2 = new Paragraph(doc);
            pSign2.Runs.Add(new Run(doc, "Lucas with lucian"));
            pSign2.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pSign2.ParagraphFormat.SpaceBefore = 16;
            pSign2.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            textBoxShapeSign.AppendChild(pSign2);

            Paragraph pSign3 = new Paragraph(doc);
            pSign3.Runs.Add(new Run(doc, "Date: 6/19/2008 7:00:00 AM"));
            pSign3.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pSign3.ParagraphFormat.SpaceBefore = 8;
            pSign3.ParagraphFormat.Style = doc.Styles["StyleNameData"];
            textBoxShapeSign.AppendChild(pSign3);

            builder.InsertNode(textBoxShapeSign);
            return builder;
        }

        public static DocumentBuilder Name(Document doc, DocumentBuilder builder, string title, string data, double marginLeft, double top)
        {
            Shape textBoxShapeTittle = new Shape(doc, ShapeType.FlowChartAlternateProcess)
            {
                Width = 150,
                Height = 52,
                Left = marginLeft,
                Top = top
            };
            textBoxShapeTittle.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeTittle.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeTittle.FillColor = Color.FromArgb(234, 239, 244);
            textBoxShapeTittle.Stroked = false;
            if(title == "Accompanied by parents/legal guardians")
            {
                textBoxShapeTittle.Height = 71;
            }
            Paragraph pTitle = new Paragraph(doc);
            pTitle.Runs.Add(new Run(doc, title));
            pTitle.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pTitle.ParagraphFormat.SpaceBefore = 3;
            pTitle.ParagraphFormat.Style = doc.Styles["Title"];

            textBoxShapeTittle.AppendChild(pTitle);

            builder.InsertNode(textBoxShapeTittle);

            Shape textBoxShapeData = new Shape(doc, ShapeType.TextBox)
            {
                Width = 330,
                Height = textBoxShapeTittle.Height,
                Left = marginLeft + 150 + 8,
                Top = top
            };
            textBoxShapeData.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeData.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeData.FillColor = Color.FromArgb(255, 255, 255);
            textBoxShapeData.Stroked = false;

            Paragraph pData = new Paragraph(doc);
            pData.Runs.Add(new Run(doc, data));
            pData.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pData.ParagraphFormat.SpaceBefore = 3;
            pData.ParagraphFormat.Style = doc.Styles["Data"];

            textBoxShapeData.AppendChild(pData);

            builder.InsertNode(textBoxShapeData);
            return builder;
        }
    }

    public static class IntelServices
    {
        public static DocumentBuilder HeaderShape(Document doc, DocumentBuilder builder, string data, double top, double marginLeft)
        {
            Shape textBoxShapeHeader = new Shape(doc, ShapeType.TextBox)
            {
                Width = 688,
                Height =51,
                Left = marginLeft,
                Top = top
            };
            textBoxShapeHeader.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeHeader.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeHeader.FillColor = Color.FromArgb(14, 37, 84);

            Paragraph pHeader = new Paragraph(doc);
            pHeader.Runs.Add(new Run(doc, data));
            pHeader.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            pHeader.ParagraphFormat.Style = doc.Styles["Title"];
            pHeader.ParagraphFormat.SpaceBefore = 10;
            textBoxShapeHeader.AppendChild(pHeader);

            builder.InsertNode(textBoxShapeHeader);
            return builder;
        }
        public static DocumentBuilder ShapeName(Document doc, DocumentBuilder builder, string label, string data, double top, double marginLeft, double heightShape)
        {
            Shape textBoxShapeTittle = new Shape(doc, ShapeType.FlowChartAlternateProcess)
            {
                Width = 149,
                Height = heightShape,
                Left = marginLeft,
                Top = top
            };
            textBoxShapeTittle.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeTittle.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeTittle.FillColor = Color.FromArgb(234, 239, 244);
            textBoxShapeTittle.Stroked = false;
            //if (title == "Accompanied by parents/legal guardians")
            //{
            //    textBoxShapeTittle.Height = 71;
            //}
            Paragraph pTitle = new Paragraph(doc);
            pTitle.Runs.Add(new Run(doc, label));
            pTitle.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pTitle.ParagraphFormat.SpaceBefore = 2;
            pTitle.ParagraphFormat.Style = doc.Styles["Label"];

            textBoxShapeTittle.AppendChild(pTitle);

            builder.InsertNode(textBoxShapeTittle);

            Shape textBoxShapeData = new Shape(doc, ShapeType.TextBox)
            {
                Width = 183,
                Height = textBoxShapeTittle.Height,
                Left = marginLeft + 150,
                Top = top
            };
            textBoxShapeData.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBoxShapeData.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBoxShapeData.FillColor = Color.FromArgb(255, 255, 255);
            textBoxShapeData.Stroked = false;

            Paragraph pData = new Paragraph(doc);
            pData.Runs.Add(new Run(doc, data));
            pData.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            pData.ParagraphFormat.SpaceBefore = 2;
            pData.ParagraphFormat.Style = doc.Styles["Data"];

            textBoxShapeData.AppendChild(pData);

            builder.InsertNode(textBoxShapeData);
            return builder;
        }
    }
}
