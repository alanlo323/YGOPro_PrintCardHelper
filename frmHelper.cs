using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout.Renderer;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Image = Xceed.Document.NET.Image;

namespace YGOPro_PrintCardHelper
{
    public partial class frmHelper : Form
    {
        public frmHelper()
        {
            InitializeComponent();
            richTextBox1.AllowDrop = true;
            richTextBox1.DragEnter += new DragEventHandler(Form1_DragEnter);
            richTextBox1.DragDrop += new DragEventHandler(Form1_DragDrop);
        }
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string path in paths)
            {
                try
                {
                    ArrayList picPaths = new ArrayList();
                    if (File.Exists(path))
                    {
                        richTextBox1.Text += path + "\n";
                        string[] lines = File.ReadAllLines(path);
                        foreach (string line in lines)
                        {
                            if (new Regex(@"^\d{1,}$").IsMatch(line))
                            {
                                FileInfo fileProps = new FileInfo(new FileInfo(path).DirectoryName);
                                while (!fileProps.Name.Equals("deck"))
                                {
                                    fileProps = new FileInfo(fileProps.Directory.FullName);
                                }
                                fileProps = new FileInfo(fileProps.Directory.FullName);
                                picPaths.Add(fileProps.FullName + "\\pics\\" + line + ".jpg");
                            }
                        }
                    }
                    FileInfo file = new FileInfo(path);
                    string outputPath = file.FullName.Replace(file.Extension, "") + ".doc";
                    file = new FileInfo(outputPath);
                    file.Directory.Create();
                    ManipulateWord(file.FullName, (string[])picPaths.ToArray(typeof(string)));
                }
                catch (Exception ex)
                {
                    richTextBox1.Text += ex.Message + "\n";
                }
            }
            richTextBox1.Text += "All done!" + "\n";
        }
        private void ManipulateWord(string dest, string[] paths)
        {
            richTextBox1.Text += "Generating to print file(.doc): " + dest + "... " + "\n";
            var doc = DocX.Create(dest);
            doc.MarginLeft = 47;
            doc.MarginRight = 47;
            doc.MarginTop = 56;
            doc.MarginBottom = 56;
            Xceed.Document.NET.Paragraph par = doc.InsertParagraph();
            foreach (string _path in paths)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(_path);
                    richTextBox1.Text += "Processing card number: " + fileInfo.Name.Replace(".jpg", "") + "... ";
                    ImageData imageData = ImageDataFactory.Create(_path);
                    int width = (int)CentimeterToPixel(5.9, imageData.GetDpiX());
                    int height = (int)CentimeterToPixel(8.6, imageData.GetDpiY());

                    string tempPath = Path.GetTempPath() + fileInfo.Name;
                    Image img = doc.AddImage(_path);
                    Picture p = img.CreatePicture();
                    //p.Width = p.Width;
                    //p.Height = p.Height;
                    double _r = 37.788578371810449574726609963548;
                    p.Width = (int)(5.9 * _r);
                    p.Height = (int)(8.6 * _r);
                    //Create a new paragraph  
                    par.AppendPicture(p);
                    richTextBox1.Text += "Success!" + "\n";
                }
                catch (Exception ex)
                {
                    richTextBox1.Text += ex.InnerException.Message + "\n";
                }
            }
            doc.Save();
            richTextBox1.Text += "Generated print file(.doc): " + dest + "\n";
            convertWordToPdf(dest, dest.Replace(".doc", ".pdf"));
        }
        private void ManipulatePdf(string dest, string[] paths)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(dest));
            iText.Layout.Document doc = new iText.Layout.Document(pdfDoc);

            iText.Layout.Element.Table table = new iText.Layout.Element.Table(3);
            table.SetMargins(0, 0, 0, 0);
            table.SetPaddings(0, 0, 0, 0);

            foreach (string _path in paths)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(_path);
                    richTextBox1.Text += "Processing card number: " + fileInfo.Name.Replace(".jpg", "") + "... ";
                    ImageData imageData = ImageDataFactory.Create(_path);
                    int width = (int)CentimeterToPixel(5.9, imageData.GetDpiX());
                    int height = (int)CentimeterToPixel(8.6, imageData.GetDpiY());
                    System.Drawing.Image image = System.Drawing.Image.FromFile(_path);
                    System.Drawing.Image resizedImage = ResizeImage(image, width, height);

                    string tempPath = Path.GetTempPath() + fileInfo.Name;
                    resizedImage.Save(tempPath, ImageFormat.Jpeg);
                    //using (System.Drawing.Image bmpInput = System.Drawing.Image.FromFile(tempPath))
                    //{
                    //    using (Bitmap bmpOutput = new Bitmap(bmpInput))
                    //    {
                    //        ImageCodecInfo encoder = GetEncoder(ImageFormat.Png);
                    //        Encoder myEncoder = Encoder.Quality;

                    //        var myEncoderParameters = new EncoderParameters(1);
                    //        var myEncoderParameter = new EncoderParameter(myEncoder, 100L);
                    //        myEncoderParameters.Param[0] = myEncoderParameter;

                    //        bmpOutput.SetResolution(207.0f, 207.0f); // Change to any dpi
                    //        bmpOutput.Save(tempPath, encoder, myEncoderParameters);
                    //    }
                    //}
                    iText.Layout.Element.Cell cell = CreateImageCell(tempPath);
                    cell.SetMargins(0, 0, 0, 0);
                    cell.SetPaddings(0, 0, 0, 0);
                    table.AddCell(cell);
                }
                catch (Exception ex)
                {
                    richTextBox1.Text += ex.Message + "\n";
                }
            }

            doc.Add(table);

            doc.Close();
        }
        private void convertWordToPdf(string sourcedocx, string targetpdf)
        {
            richTextBox1.Text += "Converting to print file(.pdf): " + "... ";
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            var wordDocument = appWord.Documents.Open(sourcedocx);
            try
            {
                wordDocument.ExportAsFixedFormat(targetpdf, WdExportFormat.wdExportFormatPDF);
                richTextBox1.Text += "Success!" + "\n";
                richTextBox1.Text += "Generated print file(.pdf): " + targetpdf + "\n";
            }
            catch (Exception ex)
            {
                richTextBox1.Text += ex.InnerException.Message + "\n";
            }
            finally
            {
                wordDocument.Close();
                appWord.Quit();
            }
        }
        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }
        /// <summary>
        /// Resize the image to the specified width and height.
        /// </summary>
        /// <param name="image">The image to resize.</param>
        /// <param name="width">The width to resize to.</param>
        /// <param name="height">The height to resize to.</param>
        /// <returns>The resized image.</returns>
        public static Bitmap ResizeImage(System.Drawing.Image image, int width, int height)
        {
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
        double PixelToCentimeter(float pixel, int dpi)
        {
            double Centimeter = pixel * 2.54d / dpi;
            return (double)Centimeter;
        }
        float CentimeterToPixel(double Centimeter, int dpi)
        {
            double pixel = Centimeter * dpi / 2.54d;
            return (float)pixel;
        }
        private static iText.Layout.Element.Cell CreateImageCell(string path)
        {
            ImageData imageData = ImageDataFactory.Create(path);
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(imageData);
            img.SetWidth(imageData.GetWidth());
            iText.Layout.Element.Cell cell = new iText.Layout.Element.Cell().Add(img);
            cell.SetBorder(iText.Layout.Borders.Border.NO_BORDER);
            return cell;
        }
        private class OverlappingImageTableRenderer : TableRenderer
        {
            private ImageData image;

            public OverlappingImageTableRenderer(iText.Layout.Element.Table modelElement, ImageData img)
                : base(modelElement)
            {
                image = img;
            }

            public override void DrawChildren(DrawContext drawContext)
            {

                // Use the coordinates of the cell in the fourth row and the second column to draw the image
                iText.Kernel.Geom.Rectangle rect = rows[0][0].GetOccupiedAreaBBox();
                base.DrawChildren(drawContext);

                drawContext.GetCanvas().AddImage(image, rect.GetLeft(), rect.GetTop() - image.GetHeight(), false);
            }

            // If renderer overflows on the next area, iText uses getNextRender() method to create a renderer for the overflow part.
            // If getNextRenderer isn't overriden, the default method will be used and thus a default rather than custom
            // renderer will be created
            public override IRenderer GetNextRenderer()
            {
                return new OverlappingImageTableRenderer((iText.Layout.Element.Table)modelElement, image);
            }
        }

        private void frmHelper_Load(object sender, EventArgs e)
        {
            richTextBox1.Text += "Drap .ydk file(s) here to generate .docx file(s)" + "\n";
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }
    }
}
