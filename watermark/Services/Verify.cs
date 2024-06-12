using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Slides;
using Aspose.Words;

namespace watermark.Services
{
    public class Verify
    {
        public async Task<string> VerifyWatermark(IFormFile file)
        {
            MemoryStream ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Seek(0, SeekOrigin.Begin);
            var extension = Path.GetExtension(file.FileName);

            MemoryStream result = new MemoryStream();
            switch (extension)
            {
                case ".doc":
                case ".docx":
                    {
                        Aspose.Words.Document doc = new Aspose.Words.Document(ms);
                        if (doc.Watermark.Type == WatermarkType.Text || doc.Watermark.Type == WatermarkType.Image)
                        {
                            return "exist watermark in this document";
                        }
                        if (doc.Watermark.Type == WatermarkType.None)
                        {
                            return "not exist watermark in this document";
                        }
                        break;
                    }
                case ".xls":
                case ".xlsx":
                    {
                        Workbook workbook = new Workbook(ms);
                        Worksheet sheet = workbook.Worksheets[0];
                        var a = sheet.Shapes.FindIndex(a => a.Name == "avepoint");
                        if (a != -1)
                        {
                            return "exist watermark in this document";
                        }
                        else
                        {
                            return "not exist watermark in this document";
                        }
                    }
                case ".ppt":
                case ".pptx":
                    {
                        using Presentation watermark = new Presentation(ms);
                        {
                            ISlide slide = watermark.Slides[0];
                            if (slide.Shapes.ToList().FindIndex(x => x.Name == "avepoint") != -1)
                            {
                                return "exist watermark in this document";
                            }
                            else
                            {
                                return "not exist watermark in this document";
                            }
                        }
                    }
                case ".pdf":
                    {
                        Aspose.Pdf.Document doc = new Aspose.Pdf.Document(ms);
                        foreach (Artifact artifact in doc.Pages[1].Artifacts)
                        {
                            if (artifact.Subtype == Artifact.ArtifactSubtype.Watermark)
                            {
                                return "exist watermark in this document";
                            }
                        }
                        return "not exist watermark in this document";
                    }
                default:
                    {
                        throw new InvalidDataException("Not support extension file");
                    }
            }
            return "Not support extension file";
        }

    }
}
