using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Windows.Forms;
using Attachment = System.Net.Mail.Attachment;

namespace CertificateMailing
{
    public partial class MainForm : Form
    {
        private string _filePath = string.Empty;

        private Point _namePoint;

        private Guid _guid = Guid.NewGuid();

        private string _font = "Calibri";

        private Size _initialCertificateSize = new Size();

        public MainForm()
        {
            InitializeComponent();

            _initialCertificateSize = pictureBox.Size;

            ReadConfig();

            listBox.Items.Add("{Name}");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Files|*.jpg;*.jpeg;*.png;";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _filePath = openFileDialog.FileName;

                    var fileStream = openFileDialog.OpenFile();

                    using StreamReader reader = new StreamReader(fileStream);
                    pictureBox.BackgroundImage = Image.FromStream(fileStream);
                    pictureBox.ClientSize =pictureBox.BackgroundImage.Size;
                    pictureBox.Enabled = true;
                }
            }

        }

        private void btnAddText_Click(object sender, EventArgs e)
        {
            string text = txtNewListItem.Text;
            listBox.Items.Add(text);
            txtNewListItem.Text = string.Empty;

            PointF firstLocation = new Point(pictureBox.BackgroundImage.Width / 2, pictureBox.BackgroundImage.Height / 2);

            Bitmap bitmap = (Bitmap)pictureBox.BackgroundImage;

            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                using (Font arialFont = new Font(_font, (int)numFontSize.Value, FontStyle.Bold))
                {
                    StringFormat format = new StringFormat(StringFormatFlags.NoWrap)
                    {
                        LineAlignment = StringAlignment.Near,
                        Alignment = StringAlignment.Near
                    };
                    RectangleF rec = new RectangleF
                    {
                        Location = _namePoint,
                        Width = (int)numFontSize.Value * text.Length * pictureBox.Size.Width / _initialCertificateSize.Width,
                        Height = (int)numFontSize.Value * 2 * pictureBox.Size.Height / _initialCertificateSize.Height
                    };
                    graphics.DrawString(text, arialFont, Brushes.Black, rec, format);
                }
            }

            pictureBox.BackgroundImage = (Image)bitmap;
            pictureBox.Refresh();
        }

        private void DrawAfterFontSizeChanged()
        {
            if (listBox.SelectedItem != null)
            {
                Bitmap bitmap = (Bitmap)Image.FromFile(_filePath);

                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    using (Font arialFont = new Font(_font, (int)numFontSize.Value, FontStyle.Bold))
                    {
                        StringFormat format = new StringFormat(StringFormatFlags.NoWrap)
                        {
                            LineAlignment = StringAlignment.Near,
                            Alignment = StringAlignment.Near
                        };
                        RectangleF rec = new RectangleF
                        {
                            Location = _namePoint,
                            Width = (int) numFontSize.Value * listBox.SelectedItem.ToString().Length * pictureBox.Size.Width / _initialCertificateSize.Width,
                            Height = (int) numFontSize.Value * 2 * pictureBox.Size.Height / _initialCertificateSize.Height
                        };
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
            var checkList = CheckList();
            if (!string.IsNullOrEmpty(checkList))
            {
                MessageBox.Show("Please check: " + checkList);

                return;
            }

            LogIt("Sending mail... ");

            List<string> lines = txtMailList.Text.Split("\r\n").ToList();

            foreach (var item in lines)
            {
                if (item.Contains(':'))
                {
                    var arr = item.Split(':');

                    string name = arr[0];
                    string mailAddress = arr[1];

                    using (Bitmap bitmap = (Bitmap)Image.FromFile(_filePath))
                    {
                        using (Graphics graphics = Graphics.FromImage(bitmap))
                        {
                            using (Font arialFont = new Font(_font, (int)numFontSize.Value, FontStyle.Bold))
                            {
                                StringFormat format = new StringFormat(StringFormatFlags.NoWrap)
                                {
                                    LineAlignment = StringAlignment.Near,
                                    Alignment = StringAlignment.Near
                                };
                                RectangleF rec = new RectangleF { Location = _namePoint, Width = (int)numFontSize.Value * name.Length * pictureBox.Size.Width / _initialCertificateSize.Width, Height = (int)numFontSize.Value * 2 *pictureBox.Size.Height / _initialCertificateSize.Height };
                                graphics.DrawString(name, arialFont, Brushes.Black, rec, format);
                            }
                        }

                        bitmap.Save(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("netcoreapp3.1", "")) + _guid.ToString() + name + ".jpeg");
                    }

                    #region Smtp
                    var smtpClient = new SmtpClient(txtSmtpServer.Text)
                    {
                        Port = Convert.ToInt32(txtSmtpPort.Text),
                        EnableSsl = checkSSL.Checked,
                        UseDefaultCredentials = false,
                    };

                    smtpClient.Credentials = new System.Net.NetworkCredential(txtUsername.Text, txtPassword.Text);

                    var mailMessage = new MailMessage
                    {
                        From = new MailAddress(txtUsername.Text),
                        Subject = txtSubject.Text,
                        Body = txtMessageBody.Text,
                        IsBodyHtml = true,
                    };
                    mailMessage.To.Add(mailAddress);

                    var attachment = new Attachment(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("netcoreapp3.1", "")) + _guid.ToString() + name + ".jpeg", MediaTypeNames.Image.Jpeg);
                    mailMessage.Attachments.Add(attachment);

                    #endregion

                    try
                    {
                        smtpClient.Send(mailMessage);
                        LogIt("Mail sent successfully to " + mailAddress);
                    }
                    catch (Exception)
                    {
                        LogIt("Mail sent error");
                    }
                }
            }

            _guid = Guid.NewGuid();
        }

        private string CheckList()
        {
            var checkThisItems = string.Empty;

            if (!SendTestMail())
                checkThisItems += "Mail Settings /";

            if (string.IsNullOrEmpty(txtMailList.Text))
                checkThisItems += "Recipient List /";


            return checkThisItems;
        }

        void ReadConfig()
        {
            foreach (string key in ConfigurationManager.AppSettings)
            {
                string appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string configFile = Path.Combine(appPath ?? string.Empty, "App.config");
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap { ExeConfigFilename = configFile };
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);

                foreach (KeyValueConfigurationElement configItem in config.AppSettings.Settings)
                {
                    var control = this.Controls.Find(configItem.Key, true).First();

                    if (control.GetType() == typeof(TextBox))
                        ((TextBox)control).Text = configItem.Value;
                    else if (control.GetType() == typeof(CheckBox))
                        ((CheckBox)control).Checked = Convert.ToBoolean(configItem.Value);
                }
            }
        }

        void SaveConfig()
        {
            string appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configFile = Path.Combine(appPath ?? string.Empty, "App.config");
            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap { ExeConfigFilename = configFile };
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);

            config.AppSettings.Settings[txtSmtpServer.Name].Value = txtSmtpServer.Text;
            config.AppSettings.Settings[txtSmtpPort.Name].Value = txtSmtpPort.Text;
            config.AppSettings.Settings[checkSSL.Name].Value = checkSSL.Checked.ToString();
            config.AppSettings.Settings[txtUsername.Name].Value = txtUsername.Text;
            config.AppSettings.Settings[txtPassword.Name].Value = txtPassword.Text;
            config.Save();
            LogIt("Config saved");
        }

        private void SendTestMail(object sender, EventArgs e)
        {
            SaveConfig();

            SendTestMail();
        }

        private bool SendTestMail()
        {
            try
            {
                var smtpClient = new SmtpClient(txtSmtpServer.Text, Convert.ToInt32(txtSmtpPort.Text));
                smtpClient.EnableSsl = checkSSL.Checked;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(txtUsername.Text,
                    txtPassword.Text);
                smtpClient.Timeout = 10000;

                var mailMessage = new MailMessage
                {
                    From = new MailAddress(txtUsername.Text),
                    Subject = "Test Mail",
                    Body = "Test Mail"
                };

                mailMessage.To.Add("gorkemhacioglu@hotmail.com");

                smtpClient.Send(mailMessage);
                LogIt("Test mail sent successfully to " + mailMessage.To.First());

                return true;

            }
            catch (Exception)
            {
                LogIt("Test mail couldn't send");
                return false;
            }
        }

        private void LogIt(string log)
        {
            txtLog.Text += log + Environment.NewLine;
        }

        private void numFontSize_ValueChanged(object sender, EventArgs e)
        {
            DrawAfterFontSizeChanged();
        }

        private void pictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (listBox.SelectedItem != null)
            {
                Bitmap bitmap = (Bitmap)Image.FromFile(_filePath);

                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    using (Font arialFont = new Font(_font, (int)numFontSize.Value, FontStyle.Bold))
                    {
                        StringFormat format = new StringFormat(StringFormatFlags.NoWrap)
                        {
                            LineAlignment = StringAlignment.Near,
                            Alignment = StringAlignment.Near
                        };
                        Point loc = new Point(e.X, e.Y);

                        _namePoint = loc;
                        RectangleF rec = new RectangleF { Location = loc, Width = (int)numFontSize.Value * 6 * pictureBox.Size.Width / _initialCertificateSize.Width, Height = (int)numFontSize.Value * 2 * pictureBox.Size.Height / _initialCertificateSize.Height };
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
    }
}
