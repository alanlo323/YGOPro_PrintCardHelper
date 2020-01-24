using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout.Renderer;
using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace YGOPro_PrintCardHelper
{
    public partial class frmHelper : Form
    {
        public frmHelper()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
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
                ArrayList picPaths = new ArrayList();
                if (File.Exists(path))
                {
                    richTextBox1.Text += path + "\n";
                    string[] lines = File.ReadAllLines(path);
                    foreach (string line in lines)
                    {
                        if (new Regex(@"^\d{1,}$").IsMatch(line))
                        {
                            richTextBox1.Text += "Processing card number: " + line + "\n";
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
                string pdfPath = file.FullName.Replace(file.Extension, "") + ".pdf";
                richTextBox1.Text += pdfPath + "\n";
                file = new FileInfo(pdfPath);
                file.Directory.Create();
                ManipulatePdf(file.FullName, (string[])picPaths.ToArray(typeof(string)));
            }
        }
        private void ManipulatePdf(string dest, string[] paths)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(dest));
            Document doc = new Document(pdfDoc);

            Table table = new Table(3);

            foreach (string _path in paths)
            {
                try
                {
                    ImageData imageData = ImageDataFactory.Create(_path);
                    int width = (int)CentimeterToPixel(5.9, imageData.GetDpiX());
                    int height = (int)CentimeterToPixel(8.6, imageData.GetDpiY());
                    System.Drawing.Image image = System.Drawing.Image.FromFile(_path);
                    System.Drawing.Image resizedImage = ResizeImage(image, width, height);
                    string tempPath = Path.GetTempPath() + "\\" + new FileInfo(_path).Name;
                    resizedImage.Save(tempPath, ImageFormat.Jpeg);
                    table.AddCell(CreateImageCell(tempPath));
                }
                catch (Exception ex)
                {
                    richTextBox1.Text += ex.Message + "\n";
                }
            }

            doc.Add(table);

            doc.Close();
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
            var destRect = new Rectangle(0, 0, width, height);
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
        private static Cell CreateImageCell(string path)
        {
            ImageData imageData = ImageDataFactory.Create(path);
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(imageData);
            img.SetWidth(imageData.GetWidth());
            Cell cell = new Cell().Add(img);
            cell.SetBorder(Border.NO_BORDER);
            return cell;
        }
        private class OverlappingImageTableRenderer : TableRenderer
        {
            private ImageData image;

            public OverlappingImageTableRenderer(Table modelElement, ImageData img)
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
                return new OverlappingImageTableRenderer((Table)modelElement, image);
            }
        }

        private void frmHelper_Load(object sender, EventArgs e)
        {
            richTextBox1.Text += "Drap ydk file(s) here to generate pdf file(s)" + "\n";
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }
    }
}
