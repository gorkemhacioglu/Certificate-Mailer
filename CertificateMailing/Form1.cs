using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificateMailing
{
    public partial class Form1 : Form
    {
        public string filePath = string.Empty;

        public Point namePoint;

        public Form1()
        {
            InitializeComponent();

            listBox.Items.Add("{Name}");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Files|*.jpg;*.jpeg;*.png;";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;

                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                        pictureBox.BackgroundImage = Image.FromStream(fileStream);
                        pictureBox.Enabled = true;
                    }
                }
            }

        }

        private void btnAddText_Click(object sender, EventArgs e)
        {
            string text = txtNewListItem.Text;
            listBox.Items.Add(text);
            txtNewListItem.Text = string.Empty;

            PointF firstLocation = new Point(pictureBox.BackgroundImage.Width/2, pictureBox.BackgroundImage.Height / 2);

            Bitmap bitmap = (Bitmap)pictureBox.BackgroundImage;

            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                using (Font arialFont = new Font("Arial", 30))
                {
                    StringFormat format = new StringFormat();
                    format.LineAlignment = StringAlignment.Center;
                    format.Alignment = StringAlignment.Near;
                    Point loc = new Point(pictureBox.Location.X + 150, (int)firstLocation.Y);
                    RectangleF rec = new RectangleF { Location = loc, Width = pictureBox.BackgroundImage.Width, Height = 35 };
                    graphics.DrawString(text, arialFont, Brushes.Black, rec, format);
                }
            }

            pictureBox.BackgroundImage = (Image)bitmap;
            pictureBox.Refresh();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (listBox.SelectedItem != null)
            {
                this.Cursor = new Cursor(Cursor.Current.Handle);

                Bitmap bitmap = (Bitmap)Image.FromFile(filePath);

                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    using (Font arialFont = new Font("Arial", 30))
                    {
                        namePoint = new Point(Cursor.Position.X, Cursor.Position.Y);
                        StringFormat format = new StringFormat();
                        format.LineAlignment = StringAlignment.Center;
                        format.Alignment = StringAlignment.Near;
                        Point loc = new Point(pictureBox.Location.X + 150, (int)namePoint.Y);
                        RectangleF rec = new RectangleF { Location = loc, Width = pictureBox.BackgroundImage.Width, Height = 35 };
                        graphics.DrawString(listBox.SelectedItem.ToString(), arialFont, Brushes.Black, rec, format);
                    }
                }

                pictureBox.BackgroundImage = (Image)bitmap;
                pictureBox.Refresh();
            }
            else 
            {
                MessageBox.Show("Please select an item to locate first");
            }
            
        }

        private void btnSendMail_Click(object sender, EventArgs e)
        {
            List<string> lines = txtMailList.Text.Split("\r\n").ToList();

            foreach (var item in lines)
            {
                if (item.Contains(':')) {
                    var arr = item.Split(':');

                    string name = arr[0];
                    string mailAddress = arr[1];

                    Bitmap bitmap = (Bitmap)Image.FromFile(filePath);

                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        using (Font arialFont = new Font("Arial", 30))
                        {
                            StringFormat format = new StringFormat();
                            format.LineAlignment = StringAlignment.Center;
                            format.Alignment = StringAlignment.Near;
                            Point loc = new Point(pictureBox.Location.X + 150, (int)namePoint.Y);
                            RectangleF rec = new RectangleF { Location = loc, Width = pictureBox.BackgroundImage.Width, Height = 35 };
                            graphics.DrawString(name, arialFont, Brushes.Black, rec, format);
                        }
                    }

                    bitmap.Save("C:\\Users\\Görkem\\Downloads\\" + mailAddress + ".jpeg");
                }
            }
        }
    }
}
