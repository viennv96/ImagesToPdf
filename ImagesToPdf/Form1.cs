using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ImagesToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            CheckForIllegalCrossThreadCalls = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FileFolderDialog sfd = new FileFolderDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Thread t = new Thread(() =>
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        ssStatus.Text = String.Format("Đang xử lý {0}!", row.Cells[1].Value.ToString());
                        try
                        {
                            List<string> images = listAllImages(row.Cells[0].Value.ToString());
                            ImagesFoldersToPdf(images, sfd.SelectedPath, row.Cells[1].Value.ToString());
                            ssStatus.Text = String.Format("Xử lý {0} hoàn tất!", row.Cells[1].Value.ToString());
                            row.DefaultCellStyle.BackColor = Color.LimeGreen;
                        }
                        catch (Exception ex)
                        {
                            ssStatus.Text = String.Format("Xử lý {0} Lỗi!", row.Cells[1].Value.ToString());
                            row.DefaultCellStyle.BackColor = Color.Red;
                            MessageBox.Show(ex.ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    MessageBox.Show(String.Format("Completed!\nAll File Save here \"{0}\"", sfd.SelectedPath), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
                t.Start();
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileFolderDialog ffd = new FileFolderDialog();
            if (ffd.ShowDialog() == DialogResult.OK)
            {
                AddDataToTable(ffd.SelectedPath);
            }
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] items = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string item in items)
            {
                AddDataToTable(item);
            }
        }

        void AddDataToTable(string folderPath)
        {
            string[] folders = Directory.GetDirectories(folderPath);
            if (folders.Length != 0)
            {
                foreach (string folder in folders)
                {
                    string outputName = new DirectoryInfo(folder).Name;
                    dataGridView1.Rows.Add(folder, string.Format("{0}.pdf", outputName));
                }
            }
            else
            {
                string outputName = new DirectoryInfo(folderPath).Name;
                dataGridView1.Rows.Add(folderPath, string.Format("{0}.pdf", outputName));
            }
        }

        List<string> listAllImages(string folderPath)
        {
            List<string> images = new List<string>();
            string[] all = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
            foreach (string record in all)
            {
                if (Directory.Exists(record))
                {
                    listAllImages(record);
                }
                else
                {
                    FileInfo fi = new FileInfo(record);
                    if (fi.Extension.ToLower() == ".png" || fi.Extension.ToLower() == ".jpg" || fi.Extension.ToLower() == ".jpeg" || fi.Extension.ToLower() == ".bmp")
                    {
                        images.Add(record);
                    }
                }
            }
            return images;
        }

        void ImagesFoldersToPdf(List<string> images, string outputFolderPath, string outputFileName)
        {
            PdfDocument ouput = new PdfDocument();

            foreach (string image in images)
            {

                XImage im = XImage.FromFile(image);

                if (File.Exists("temp.png"))
                {
                    File.Delete("temp.png");
                }

                var bitmap = Bitmap.FromFile(image);

                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        break;
                    case 1:
                        bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        break;
                    case 2:
                        bitmap.RotateFlip(RotateFlipType.Rotate180FlipNone);
                        break;
                    case 3:
                        bitmap.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        break;
                }


                if (im.Width > im.Height)
                {
                    bitmap.RotateFlip(RotateFlipType.Rotate270FlipNone);
                }
                bitmap.Save("temp.png", ImageFormat.Png);
                im.Dispose();
                bitmap.Dispose();

                XImage img = XImage.FromFile("temp.png");

                // each source file saeparate
                PdfDocument doc = new PdfDocument();
                PdfPage p = new PdfPage();

                double width = img.Width;
                double height = img.Height;
                double w = 0, h = 0;
                double rate = width / height;
                double x = 0, y = 0;

                if (rate < (p.Width / p.Height))
                {
                    //p.Orientation = PdfSharp.PageOrientation.Portrait;
                    h = p.Height;
                    w = rate * h;
                    x = (p.Width - w) / 2;
                }
                else
                {
                    //p.Orientation = PdfSharp.PageOrientation.Landscape;
                    w = p.Width;
                    h = w / rate;
                    y = (p.Height - h) / 2;
                }

                doc.Pages.Add(p);
                XGraphics xgr = XGraphics.FromPdfPage(doc.Pages[0]);
                xgr.DrawImage(img, x, y, w, h);
                img.Dispose();
                xgr.Dispose();
                //  save to destination file
                FileInfo fi = new FileInfo("temp.png");
                doc.Save(fi.FullName.Replace(fi.Extension, ".PDF"));

                PdfDocument inputDocument = PdfReader.Open(fi.FullName.Replace(fi.Extension, ".PDF"), PdfDocumentOpenMode.Import);
                PdfPage page = inputDocument.Pages[0];
                ouput.AddPage(page);
                File.Delete(fi.FullName.Replace(fi.Extension, ".PDF"));
                doc.Close();
                doc.Dispose();
            }
            ouput.Save(string.Format("{0}\\{1}", outputFolderPath, outputFileName));
            ouput.Close();
            ouput.Dispose();
            GC.Collect();
        }
    }
}