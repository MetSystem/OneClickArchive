using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Reflection;
using iTextSharp.text.io;

namespace Met.Report
{
    public class ReportManager
    {
        private static BaseFont _baseFont;

        public static BaseFont BaseFont
        {
            get
            {
                if (!File.Exists("SIMYOU.TTF"))
                {
                    Assembly myAssembly = Assembly.GetExecutingAssembly();
                    Stream stream = myAssembly.GetManifestResourceStream("Met.Report.SIMYOU.TTF");

                    byte[] bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    stream.Seek(0, SeekOrigin.Begin);

                    FileStream fs = new FileStream("SIMYOU.TTF", FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(bytes);
                    bw.Close();
                    fs.Close();
                }

                if (_baseFont == null)
                {
                    StreamUtil.AddToResourceSearch("iTextAsian.dll");
                    StreamUtil.AddToResourceSearch("iTextAsianCmaps.dll");
                    _baseFont = BaseFont.CreateFont("SIMYOU.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                }

                return _baseFont;
            }
        }

        private static ReportManager _ReportManager;
        public static ReportManager Instace
        {
            get
            {
                if (_ReportManager == null)
                {
                    _ReportManager = new ReportManager();
                }
                return _ReportManager;
            }
        }

        private Font GetFont()
        {
            return new Font(BaseFont);
        }

        public void BuildReport(string[] fileList, string targetFile)
        {
            PdfReader reader;
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(targetFile, FileMode.Create));
            document.Open();

            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage newPage;

            for (int i = 0; i < fileList.Length; i++)
            {
                reader = new PdfReader(fileList[i]);
                int iPageNum = reader.NumberOfPages;

                for (int j = 1; j <= iPageNum; j++)
                {
                    newPage = writer.GetImportedPage(reader, j);
                    document.NewPage();
                    cb.AddTemplate(newPage, 0, 0);
                }
            }
            document.Close();
        }

        public void BuildReport(ReportInfo report, string targetFile)
        {
            Font font = GetFont();
            Document document = new Document(PageSize.A4);

            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(targetFile, FileMode.Create));
            bool hasCoverPath = !string.IsNullOrEmpty(report.CoverPath);
            // Our custom Header and Footer is done using Event Handler
            MyPageEventHelper PageEventHandler = new MyPageEventHelper();
            writer.PageEvent = PageEventHandler;
            // Define the page header 
            PageEventHandler.BaseFont = BaseFont;
            PageEventHandler.Title = report.PageTitle;
            PageEventHandler.HeaderFont = new Font(BaseFont, 10, Font.BOLD);
            PageEventHandler.HasCover = hasCoverPath;

            //MyPdfPageEventHelpPageNo pageeventhandler = new MyPdfPageEventHelpPageNo();
            //writer.PageEvent = pageeventhandler;

            if (string.IsNullOrEmpty(report.WaterMarkName))
            {
                if (!string.IsNullOrEmpty(report.Password))
                {
                    writer.SetEncryption(PdfWriter.STRENGTH40BITS, report.Password, report.Password, PdfWriter.AllowPrinting);
                }
            }

            document.AddAuthor(report.Author);
            document.AddCreationDate();
            document.AddCreator(report.Creator);
            document.AddSubject(report.Subject);
            document.AddTitle(report.Title);
            document.AddKeywords(report.Keywords);
            document.AddHeader("Expires", "0");
            document.Open();

            PdfContentByte cb = writer.DirectContent;

            if (hasCoverPath)
            { 
                PdfReader reader = new PdfReader(report.CoverPath);
                PdfImportedPage newPage = writer.GetImportedPage(reader, 1);
                cb.AddTemplate(newPage, 0, 0);
            }

            // 目录导航
            int chapterNum = 1;

            Chapter chapter0 = new Chapter(new Paragraph(report.Title, new Font(BaseFont, 15f, Font.BOLD)), chapterNum);
            List list = new List(List.UNORDERED, false, 10f);
            list.ListSymbol = new Chunk("\u2022", font);
            list.IndentationLeft = 15f;
            BuildList(report.PeportFiles, list, "1.", 1);
            chapter0.Add(list);
            document.Add(chapter0);

            // 生成章节   
            BuildChapter(report, document, writer, cb, font, chapter0, chapterNum);
            document.Close();
        }

        private void BuildList(List<ReportFileInfo> reportFiles, List list, string prefix, int deepin)
        {
            var font = new Font(BaseFont);
            if (!(reportFiles == null || reportFiles.Count == 0))
            {
                for (int i = 0; i < reportFiles.Count; i++)
                {
                    var subPrefix = string.Format("{0}{1}.", prefix, (i + 1));
                    var reportInfo = reportFiles[i];
                    list.Add(new ListItem(string.Format("{0} {1}", subPrefix, reportInfo.Name), font));
                    if (!(reportInfo == null || reportInfo.SubFiles == null || reportInfo.SubFiles.Count == 0))
                    {
                        BuildList(reportInfo.SubFiles, list, subPrefix, 1);
                    }
                }
            }
        }

        /// <summary>
        /// Build Chapter
        /// </summary> 
        private void BuildChapter(ReportInfo report, Document document, PdfWriter writer, PdfContentByte cb, Font font, Chapter chapter, int chapterNum)
        {
            var sectionNum = chapterNum + 1;
            foreach (ReportFileInfo info in report.PeportFiles)
            {
                var section = chapter.AddSection(20f, new Paragraph(info.Name, GetFont()), sectionNum);
                BuildReportFile(document, writer, cb, info, section);

                if (!(info.SubFiles == null || info.SubFiles.Count == 0))
                {
                    BuildSection(document, writer, cb, font, info.SubFiles, section, sectionNum);
                }
            }
        }

        /// <summary>
        /// Build Section
        /// </summary> 
        private void BuildSection(Document document, PdfWriter writer, PdfContentByte cb, Font font, List<ReportFileInfo> info, Section section, int sectionNum)
        {
            var subNum = sectionNum + 1;
            foreach (var subFile in info)
            {
                var subSection = section.AddSection(20f, new Paragraph(subFile.Name, GetFont()), subNum);
                BuildReportFile(document, writer, cb, subFile, subSection);
                if (!(subFile.SubFiles == null || subFile.SubFiles.Count == 0))
                {
                    BuildSection(document, writer, cb, font, subFile.SubFiles, subSection, subNum);
                }
            }
        }

        /// <summary>
        /// Build Report File
        /// </summary> 
        private void BuildReportFile(Document document, PdfWriter writer, PdfContentByte cb, ReportFileInfo reportFile, Section section)
        {
            PdfReader reader = new PdfReader(reportFile.Path);
            int PageNum = reader.NumberOfPages;

            for (int j = 1; j <= PageNum; j++)
            {
                document.NewPage();
                PdfImportedPage newPage = writer.GetImportedPage(reader, j);
                Rectangle psize = reader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;

                if (j == 1)
                {
                    document.Add(section);
                    cb.AddTemplate(newPage, 0, -10);
                }
                else
                {
                    cb.AddTemplate(newPage, 0, 0);
                }
            }
        }

        public bool SetWatermarkByImage(string inputfilepath, string outputfilepath, string ModelPicName, float top, float left, bool HasCover = false)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                int numberOfPages = pdfReader.NumberOfPages;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                PdfContentByte waterMarkContent;
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(ModelPicName);
                image.GrayFill = 20;

                if (left < 0)
                {
                    left = width / 2 - image.Width + left;
                }

                image.SetAbsolutePosition(left, (height / 2 - image.Height) - top);
                for (int i = 1; i <= numberOfPages; i++)
                {
                    if (HasCover && i == 1)
                    {
                        continue;
                    }

                    waterMarkContent = pdfStamper.GetOverContent(i);
                    waterMarkContent.AddImage(image);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;

            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

        public bool SetWatermarkByImage(string inputfilepath, string outputfilepath, string ModelPicName, float top, float left, string userPassWord, string ownerPassWord, int permission, bool HasCover = false)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);

                int numberOfPages = pdfReader.NumberOfPages;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);

                float width = psize.Width;
                float height = psize.Height;

                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                pdfStamper.SetEncryption(false, userPassWord, ownerPassWord, permission);

                PdfContentByte waterMarkContent;
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(ModelPicName);

                image.GrayFill = 20;

                if (left < 0)
                {
                    left = width / 2 - image.Width + left;
                }

                image.SetAbsolutePosition(left, (height / 2 - image.Height) - top);

                for (int i = 1; i <= numberOfPages; i++)
                {
                    if (HasCover && i == 1)
                    {
                        continue;
                    }

                    waterMarkContent = pdfStamper.GetOverContent(i);
                    waterMarkContent.AddImage(image);
                }

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

        public void SetWatermarkByWord(string inputfilepath, string outputfilepath, string waterMarkName, bool HasCover = false)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);

                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont;
                PdfGState gs = new PdfGState();
                for (int i = 1; i < total; i++)
                {
                    if (HasCover && i == 1)
                    {
                        continue;
                    }

                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                                                           //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                                                           //透明度
                    gs.FillOpacity = 0.3f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(BaseColor.LIGHT_GRAY);
                    content.SetFontAndSize(font, 50);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 - 50, height / 2 - 50, 55);
                    //content.SetColorFill(BaseColor.BLACK);
                    //content.SetFontAndSize(font, 8);
                    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

        public void SetWatermarkByWord(string inputfilepath, string outputfilepath, string waterMarkName, string userPassWord, string ownerPassWord, int permission, bool HasCover = false)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                // 设置密码   
                pdfStamper.SetEncryption(false, userPassWord, ownerPassWord, permission);

                int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont;
                PdfGState gs = new PdfGState();

                for (int i = 1; i < total; i++)
                {
                    if (HasCover && i == 1)
                    {
                        continue;
                    }

                    content = pdfStamper.GetOverContent(i);
                    var len = Encoding.UTF8.GetByteCount(waterMarkName);
                    gs.FillOpacity = 0.3f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);

                    content.BeginText();
                    content.SetColorFill(BaseColor.LIGHT_GRAY);
                    var fontSize = 50;
                    if (len > 40 && len <= 60)
                    {
                        fontSize = 35;
                    }
                    else
                    {
                        fontSize = 30;
                    }
                    content.SetFontAndSize(font, fontSize);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2, height / 2 + (len), 55);
                    //content.SetColorFill(BaseColor.BLACK);
                    //content.SetFontAndSize(font, 8);
                    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                    content.EndText();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

    }
}
