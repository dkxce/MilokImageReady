using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using System.Xml;
using System.Xml.Serialization;

namespace MilokImageReady
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void listBoxControl1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void listBoxControl1_DragDrop(object sender, DragEventArgs e)
        {
            AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void listBoxControl1_KeyDown(object sender, KeyEventArgs e)
        {            
            if (e.KeyCode == Keys.Delete)
            {
                List<int> lint = new List<int>();
                foreach(int index in listBox1.SelectedIndices) lint.Add(index);
                lint.Sort();
                for(int i=lint.Count-1;i>=0;i--)
                    listBox1.Items.RemoveAt(lint[i]);
                UpdateStatus("Выбрано файлов: " + listBox1.Items.Count.ToString());
            };
            if (e.KeyCode == Keys.Insert)
            {
                OpenFileDialog ofg = new OpenFileDialog();
                ofg.DefaultExt = ".jpg";
                ofg.Filter = "JPEG Files (*.jpg)|*.jpg";
                ofg.Multiselect = true;
                if (ofg.ShowDialog() == DialogResult.OK)
                    AddFiles(ofg.FileNames);
                ofg.Dispose();
            };
        }

        public void UpdateStatus(string status)
        {
            label1.Text = status;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            PardesiServices.WinControls.FolderBrowser fb = new PardesiServices.WinControls.FolderBrowser();
            if (fb.ShowDialog() == DialogResult.OK)
                AddFiles(System.IO.Directory.GetFiles(fb.DirectoryPath));
            fb.Dispose();
            return;
        }

        private void AddFiles(string[] files)
        {
            foreach (string file in files)
                if (System.IO.Path.GetExtension(file).ToLower() == ".jpg")
                {
                    bool ex = false;
                    foreach (object obj in listBox1.Items)
                        if (((FileDropped)obj).fileName == file) ex = true;
                    if (!ex) listBox1.Items.Add(new FileDropped(file));
                };
            UpdateStatus("Files To Process: " + listBox1.Items.Count.ToString());
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            listBoxControl1_KeyDown(this,new KeyEventArgs(Keys.Insert));
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            listBoxControl1_KeyDown(this, new KeyEventArgs(Keys.Delete));
        }


        List<IExifRenamer> exifRenemers = new List<IExifRenamer>();
        private void MainForm_Load(object sender, EventArgs e)
        {
            string[] ExifKeys = XmlSaved<string[]>.Load(XmlSaved<int>.GetCurrentDir() + @"\ExifKeys.xml");
            keyDataGridViewTextBoxColumn.Items.AddRange(ExifKeys);

            foreach (string dll in System.IO.Directory.GetFiles(XmlSaved<int>.GetCurrentDir(), "*.dll"))
            {
                try
                {
                    IExifRenamer exr = XmlSaved<IExifRenamer>.LoadFromDLL(dll);
                    exifRenemers.Add(exr);
                    RenamerInterface.Items.Add(exr.GetTitle());
                }
                catch { };
            };

            RenamerInterface.SelectedIndex = 0;
            SortAfter.SelectedIndex = 1;
            AddFileNo.SelectedIndex = 2;
            WFontStyle.SelectedIndex = 1;
            WatermarkStyle.SelectedIndex = 0;
            WatermarkOverwrite.SelectedIndex = 0;
            AfterRename.SelectedIndex = 1;
            AfterWatermarked.SelectedIndex = 1;
            ResizeType.SelectedIndex = 0;
            AspectRatio.SelectedIndex = 4;
            ResizeOverwrite.SelectedIndex = 0;
            AfterResize.SelectedIndex = 1;

            // Load Defaults
            string fileName = XmlSaved<int>.GetCurrentDir()+@"\default.mirExif";
            if(System.IO.File.Exists(fileName))
                try
                {
                    exifValueItmBindingSource.Clear();
                    ExifValueItm[] list = XmlSaved<ExifValueItm[]>.Load(fileName);
                    foreach (ExifValueItm itm in list) exifValueItmBindingSource.Add(itm);
                }
                catch {};

            fileName = XmlSaved<int>.GetCurrentDir()+@"\default.mirWtmk";
            if(System.IO.File.Exists(fileName))
                try
                {
                    WatermarkCfg w = XmlSaved<WatermarkCfg>.Load(fileName);
                    Image2Font.Text = w.Image2Font.ToString().Replace(",",".");
                    ImageOpacity.Text = w.ImageOpacity.ToString();
                    ImageQuality.Text = w.ImageQuality.ToString();
                    WatermarkText.Text = w.Text;                    
                    WatermarkImage.Text = w.Image;
                    TextOpacity.Text = w.Opacity.ToString();
                    WFontName.Text = w.FontName;
                    WFontStyle.SelectedIndex = w.FontStyle;                    
                    WatermarkStyle.SelectedIndex = w.WatermarkStyle;
                }
                catch { };
            fileName = XmlSaved<int>.GetCurrentDir() + @"\presets.cfg";
            if(System.IO.File.Exists(fileName))
                try
                {
                    SaveCFG svg = XmlSaved<SaveCFG>.Load(fileName);
                    try
                    {
                        RenamerInterface.SelectedIndex = svg.RenamerInterface;
                    }
                    catch { };
                    SortAfter.SelectedIndex = svg.SortAfter;
                    AddFileNo.SelectedIndex = svg.AddFileNo;
                    AfterRename.SelectedIndex = svg.AfterRename;

                    WatermarkOverwrite.SelectedIndex = svg.WatermarkOverwrite;
                    AfterWatermarked.SelectedIndex = svg.AfterWatermarked;

                    ResizeType.SelectedIndex = svg.ResizeType;
                    ResizeHeight.Text = svg.ResizeHeight.ToString();
                    ResizeWidth.Text = svg.ResizeWidth.ToString();
                    AspectRatio.SelectedIndex = svg.AspectRatio;
                    ResizeQuality.Text = svg.ResizeQuality.ToString();
                    EraseQuality.Text = svg.EraseQuality.ToString();
                    ResizeOverwrite.SelectedIndex = svg.ResizeOVerwrite;
                    AfterResize.SelectedIndex = svg.AfterResize;
                }
                catch { };

            System.IO.Directory.CreateDirectory(XmlSaved<int>.GetCurrentDir() + @"\Configurations");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (exifValueItmBindingSource.Count == 0) return;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = XmlSaved<int>.GetCurrentDir() + @"\Configurations";
            sfd.DefaultExt = ".mirExif";
            sfd.Filter = "MilokImageReady Exif Cfg (*.mirExif)|*.mirExif";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                List<ExifValueItm> list = new List<ExifValueItm>();
                foreach(object obj in exifValueItmBindingSource) list.Add((ExifValueItm)obj);
                XmlSaved<ExifValueItm[]>.Save(sfd.FileName, list.ToArray());
            };
            sfd.Dispose();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = XmlSaved<int>.GetCurrentDir() + @"\Configurations";
            ofd.DefaultExt = ".mirExif";
            ofd.Filter = "MilokImageReady Exif Cfg (*.mirExif)|*.mirExif";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                exifValueItmBindingSource.Clear();
                ExifValueItm[] list = XmlSaved<ExifValueItm[]>.Load(ofd.FileName);
                foreach (ExifValueItm itm in list) exifValueItmBindingSource.Add(itm);
            };
            ofd.Dispose();
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string fileName = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            if (System.IO.Path.GetExtension(fileName).ToLower() != ".mirexif") return;
            exifValueItmBindingSource.Clear();
            ExifValueItm[] list = XmlSaved<ExifValueItm[]>.Load(fileName);
            foreach (ExifValueItm itm in list)
            {
                if (!keyDataGridViewTextBoxColumn.Items.Contains(itm.Key)) keyDataGridViewTextBoxColumn.Items.Add(itm.Key);
                exifValueItmBindingSource.Add(itm);
            };
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (IExifRenamer er in exifRenemers)
                if (er.GetTitle() == RenamerInterface.Text)
                {
                    string[] keys = er.GetExifKeys();
                    string txt = keys[0];
                    for(int i=1;i<keys.Length;i++) txt+=", "+keys[i];
                    label4.Text = txt;
                };
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;
            if (exifValueItmBindingSource.Count == 0) return;
            if (MessageBox.Show("Do you really want to rewrite Exif data for "+listBox1.Items.Count.ToString()+" images?", "Exif Rewrite", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            tabControl1.SelectedIndex = 6;
            listBox2.Items.Clear();
            Console("Exif Rewrite",true);
            Console("Files to Process: " + listBox1.Items.Count.ToString(), true);
            Console("Begin", true);
            SetProgress(0, listBox1.Items.Count);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                FileDropped fd = (FileDropped)listBox1.Items[i];
                Console(" WriteExif '" + fd.ToString() + "' .. ", true);
                using (Exiv2Net.Image image = new Exiv2Net.Image(fd.fileName))
                {
                    //image["Exif.Photo.DateTimeOriginal"]
                    foreach(ExifValueItm val in exifValueItmBindingSource)
                    {
                        image[val.Key] = new Exiv2Net.AsciiString(val.Value);
                    };
                    image.Save();
                    image.Dispose();
                };
                Console("OK",false);
                SetProgress(i+1, listBox1.Items.Count);
            };
            Console("Total Files: " + listBox1.Items.Count.ToString(), true);
            Console("End",true);

            SaveDefsExifRewrite();
        }

        public void Console(string text, bool newLine)
        {
            if (newLine) listBox2.Items.Insert(0, "");
            listBox2.Items[0] += text;
            listBox2.Refresh();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;
            if (MessageBox.Show("Do you really want to rename images from Exif data for " + listBox1.Items.Count.ToString() + " files?", "Exif Rename", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            tabControl1.SelectedIndex = 6;
            listBox2.Items.Clear();
            Console("Exif Rename", true);
            Console("Files to Process: " + listBox1.Items.Count.ToString(), true);
            Console("Begin", true);

            List<fileRN> filesNew = new List<fileRN>();
            SetProgress(0, listBox1.Items.Count);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                FileDropped fd = (FileDropped)listBox1.Items[i];
                Console(" ReadExif '" + fd.ToString() + "' .. ", true);
                using (Exiv2Net.Image image = new Exiv2Net.Image(fd.fileName))
                {
                    string ext = System.IO.Path.GetExtension(fd.fileName).ToLower();
                    string namOld = System.IO.Path.GetFileNameWithoutExtension(fd.fileName);
                    string newName = "";

                    IExifRenamer ren = exifRenemers[RenamerInterface.SelectedIndex];
                    List<string> keyvalues = new List<string>();
                    foreach (string key in ren.GetExifKeys())
                    {
                        try { keyvalues.Add(image[key].ToString()); }
                        catch { keyvalues.Add(String.Empty); };
                    };
                    newName = ren.GetNewNameFromExifData(keyvalues.ToArray());
                    if((newName.Length > 0) && (namOld != newName)) filesNew.Add(new fileRN(fd.fileName, namOld, newName));
                    image.Dispose();
                };
                Console("OK", false);
                SetProgress(i+1, listBox1.Items.Count);
            };

            if (SortAfter.SelectedIndex > 0) filesNew.Sort(filesNew[0]);
            if (SortAfter.SelectedIndex == 2)
            {
                fileRN[] fdd = filesNew.ToArray();
                filesNew.Clear();
                for (int i = fdd.Length - 1; i >= 0; i--) filesNew.Add(fdd[i]);
            };

            for (int i = 0; i < filesNew.Count; i++)
            {
                string nn = filesNew[i].newName;
                if (AddFileNo.SelectedIndex == 1) nn = IntToStr(i + 1, filesNew.Count.ToString().Length) + " - " + nn;
                if (AddFileNo.SelectedIndex == 2) nn += " - " + IntToStr(i + 1, filesNew.Count.ToString().Length);
                string nnn = nn + "";
                if (filesNew[i].oldName != nn)
                {
                    int tyies = 1;
                    while (System.IO.File.Exists(filesNew[i].fullPath.Replace(filesNew[i].oldName, nnn)))
                        nnn = nn + ", " + (tyies++).ToString();
                    System.IO.File.Move(filesNew[i].fullPath, filesNew[i].fullPath.Replace(filesNew[i].oldName, nnn));
                };

                if(AfterRename.SelectedIndex == 1)
                for (int x = 0; x < listBox1.Items.Count; x++)
                {
                    FileDropped fd = (FileDropped)listBox1.Items[x];
                    if (fd.fileName == filesNew[i].fullPath)
                    {
                        fd.fileName = filesNew[i].fullPath.Replace(filesNew[i].oldName, nnn);
                        listBox1.Items[x] = fd;
                    };
                };
                Console("  Rename '" + nnn + "' <-- '" + filesNew[i].oldName + "'",true);
            };
            

            Console("Total Files: " + listBox1.Items.Count.ToString(), true);
            Console("End", true);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = XmlSaved<int>.GetCurrentDir() + @"\Configurations";
            ofd.DefaultExt = ".mirWtmk";
            ofd.Filter = "MilokImageReady Watermark Cfg (*.mirWtmk)|*.mirWtmk";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                WatermarkCfg w = XmlSaved<WatermarkCfg>.Load(ofd.FileName);
                Image2Font.Text = w.Image2Font.ToString();
                ImageOpacity.Text = w.ImageOpacity.ToString();
                ImageQuality.Text = w.ImageQuality.ToString();
                WatermarkText.Text = w.Text;
                WatermarkImage.Text = w.Image;
                TextOpacity.Text = w.Opacity.ToString();
                WFontName.Text = w.FontName;
                WFontStyle.SelectedIndex = w.FontStyle;
                WatermarkStyle.SelectedIndex = w.WatermarkStyle;
            };
            ofd.Dispose();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (exifValueItmBindingSource.Count == 0) return;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = XmlSaved<int>.GetCurrentDir() + @"\Configurations";
            sfd.DefaultExt = ".mirWtmk";
            sfd.Filter = "MilokImageReady Watermark Cfg (*.mirWtmk)|*.mirWtmk";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                WatermarkCfg w = new WatermarkCfg();
                w.Image2Font = ParseDouble(Image2Font.Text);
                w.ImageOpacity = Convert.ToInt32(ImageOpacity.Text.Trim());
                w.ImageQuality = Convert.ToInt32(ImageQuality.Text.Trim());
                w.Text = WatermarkText.Text;
                w.Image = WatermarkImage.Text;
                w.Opacity = Convert.ToInt32(TextOpacity.Text.Trim());
                w.FontName = WFontName.Text;
                w.FontStyle = WFontStyle.SelectedIndex;
                w.WatermarkStyle = WatermarkStyle.SelectedIndex;
                XmlSaved<WatermarkCfg>.Save(sfd.FileName, w);
            };
            sfd.Dispose();
        }

        private void SaveDefsExifRewrite()
        {
            List<ExifValueItm> list = new List<ExifValueItm>();
            foreach (object obj in exifValueItmBindingSource) list.Add((ExifValueItm)obj);
            XmlSaved<ExifValueItm[]>.Save(XmlSaved<int>.GetCurrentDir() + @"\default.mirExif", list.ToArray());
        }

        private void SaveDefsWtmkRewrite()
        {
            WatermarkCfg w = new WatermarkCfg();
            w.Image2Font = ParseDouble(Image2Font.Text);
            w.ImageOpacity = Convert.ToInt32(ImageOpacity.Text.Trim());
            w.ImageQuality = Convert.ToInt32(ImageQuality.Text.Trim());
            w.Text = WatermarkText.Text;
            w.Image = WatermarkImage.Text;
            w.Opacity = Convert.ToInt32(TextOpacity.Text.Trim());
            w.FontName = WFontName.Text;
            w.FontStyle = WFontStyle.SelectedIndex;
            w.WatermarkStyle = WatermarkStyle.SelectedIndex;
            XmlSaved<WatermarkCfg>.Save(XmlSaved<int>.GetCurrentDir() + @"\default.mirWtmk", w);
        }

        public void SetProgress(int curValue, int MaxValue)
        {
            StatusText.Text = "Progress: " + curValue.ToString() + "/" + MaxValue.ToString() + " - "+((double)curValue/(double)MaxValue*100).ToString("0.00").Replace(",",".")+"%";
            if (StatusVal.Maximum != MaxValue) StatusVal.Maximum = MaxValue;
            StatusVal.Value = curValue;
            StatusText.Refresh();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string watermarkFileName = WatermarkImage.Text;
            if (watermarkFileName.IndexOf(":") < 0) watermarkFileName = XmlSaved<int>.GetCurrentDir() + @"\" + watermarkFileName;

            if (listBox1.Items.Count == 0) return;
            if (exifValueItmBindingSource.Count == 0) return;
            if (WatermarkStyle.SelectedIndex != 0)
                if (!System.IO.File.Exists(watermarkFileName))
                {
                    MessageBox.Show("Watermark File Not Found!", "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                };
            if (MessageBox.Show("Do you really want to add watermark in " + listBox1.Items.Count.ToString() + " images?", "Watermark", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            tabControl1.SelectedIndex = 6;
            listBox2.Items.Clear();
            Console("Add Watermark", true);
            Console("Files to Process: " + listBox1.Items.Count.ToString(), true);
            Console("Begin", true);

            System.Drawing.Image wm = null;
            double cropFactor = 0;
            if (WatermarkStyle.SelectedIndex != 0)
            {
                wm = System.Drawing.Image.FromFile(watermarkFileName);
                cropFactor = (double)wm.Width / (double)wm.Height;
                if (ImageOpacity.Text.Trim() != "100") wm = SetImgOpacity(wm, (float)(ParseDouble(ImageOpacity.Text.Trim()) / 100));
            };

            SetProgress(0, listBox1.Items.Count);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                FileDropped fd = (FileDropped)listBox1.Items[i];
                Console(" Watermark '" + fd.ToString() + "' .. ", true);
                //
                string tmpf = fd.fileName + ".tmp";
                System.Drawing.Image im = System.Drawing.Image.FromFile(fd.fileName);
                System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(im);

                int ldiff = 0;
                int fontHeight = (int)(Math.Max(im.Height, im.Width) / 60);
                if (fontHeight < 12) fontHeight = 12;

                if (WatermarkStyle.SelectedIndex != 1)
                {
                    FontStyle fs = FontStyle.Regular;
                    if (WFontStyle.SelectedIndex == 1) fs = FontStyle.Bold;
                    if (WFontStyle.SelectedIndex == 2) fs = FontStyle.Bold;
                    if (WFontStyle.SelectedIndex == 3) fs = FontStyle.Italic | FontStyle.Bold;
                    System.Drawing.Font fnt = new System.Drawing.Font(WFontName.Text.Trim(), fontHeight, fs, GraphicsUnit.Pixel);
                    System.Drawing.SizeF sf = g.MeasureString(WatermarkText.Text, fnt);
                    System.Drawing.SolidBrush brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb((int)(ParseDouble(TextOpacity.Text.Trim())/100*255), System.Drawing.Color.White));

                    if (WatermarkStyle.SelectedIndex != 0)
                    {                     
                        double wmwi = (double)fnt.Height * ParseDouble(Image2Font.Text);
                        double wmhe = wmwi/cropFactor;
                        if (WatermarkStyle.SelectedIndex == 2)
                        {
                            ldiff = (int)wmwi + 10;
                            g.DrawImage(wm, (float)(im.Width - wmwi - 5), (float)(im.Height - wmhe - 5), (float)wmwi, (float)wmhe);
                        }
                        else
                        {
                            g.DrawImage(wm, (float)(im.Width - sf.Width - 10 - wmwi), (float)(im.Height - wmhe - 5), (float)wmwi, (float)wmhe);
                        };
                    };

                    g.DrawString(WatermarkText.Text, fnt, brush, im.Width - sf.Width - 6 - ldiff, im.Height - sf.Height - 6);
                    brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(Convert.ToInt32(TextOpacity.Text.Trim()), System.Drawing.Color.Black));
                    g.DrawString(WatermarkText.Text, fnt, brush, im.Width - sf.Width - 5 - ldiff, im.Height - sf.Height - 5);                    
                }
                else
                {
                    double wmwi = fontHeight * ParseDouble(Image2Font.Text);
                    double wmhe = wmwi / cropFactor;
                    g.DrawImage(wm, (float)(im.Width - wmwi - 5), (float)(im.Height - wmhe - 5), (float)wmwi, (float)wmhe);
                };                
                string subdir = System.IO.Path.GetDirectoryName(fd.fileName) + @"\Watermarked\";
                if (!System.IO.Directory.Exists(subdir)) System.IO.Directory.CreateDirectory(subdir);
                string savedFN = subdir + System.IO.Path.GetFileName(fd.fileName);
                System.Drawing.Imaging.EncoderParameters pars = new System.Drawing.Imaging.EncoderParameters(1);
                pars.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, Convert.ToInt32(ImageQuality.Text.Trim()));
                System.Drawing.Imaging.ImageCodecInfo ici = getEncoderInfo("image/jpeg");
                im.Save(savedFN, ici,pars);
                g.Dispose();
                im.Dispose();
                if (WatermarkOverwrite.SelectedIndex == 1)
                {
                    System.IO.File.Delete(fd.fileName);
                    System.IO.File.Move(subdir + System.IO.Path.GetFileName(fd.fileName), fd.fileName);
                };
                //
                Console("OK", false);
                SetProgress(i+1, listBox1.Items.Count);

                if (AfterWatermarked.SelectedIndex == 1)
                    for (int x = 0; x < listBox1.Items.Count; x++)
                    {
                        FileDropped f2d = (FileDropped)listBox1.Items[x];
                        if (f2d.fileName == fd.fileName)
                        {
                            f2d.fileName = savedFN;
                            f2d.subdir += @"Watermarked\";
                            listBox1.Items[x] = f2d;
                        };
                    };
            };
            if (WatermarkStyle.SelectedIndex != 0) wm.Dispose();
            Console("Total Files: " + listBox1.Items.Count.ToString(), true);
            Console("End", true);

            SaveDefsWtmkRewrite();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".jpg";
            ofd.Filter = "Images (*.jpg;*.png;*.gif;*.bmp)|*.jpg;*.png;*.gif;*.bmp";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                WatermarkImage.Text = ofd.FileName;
            };
            ofd.Dispose();
        }

        #region STATIC
        private static System.Drawing.Imaging.ImageCodecInfo getEncoderInfo(string mimeType)
        {
            // Get image codecs for all image formats
            System.Drawing.Imaging.ImageCodecInfo[] codecs = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();

            // Find the correct image codec
            for (int i = 0; i < codecs.Length; i++)
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];
            return null;
        }

        private static string IntToStr(int value, int digits)
        {
            int dgts = digits > 2 ? digits : 4;
            string res = value.ToString();
            while (res.Length < dgts) res = "0" + res;
            return res;
        }

        private static double ParseDouble(string StrDouble)
        {
            //TryParse not in CF.NET
            //double result;
            //Double.TryParse(StrDouble, System.Globalization.NumberStyles.Any, null, out result);
            //return result;

            char[] ca = StrDouble.ToCharArray();
            foreach (char c in ca) //StrDouble
            {
                if (!Char.IsDigit(c) & !c.Equals('.') & !c.Equals('-'))
                    return 0;
            };
            System.Globalization.CultureInfo ci = System.Globalization.CultureInfo.CurrentCulture;
            System.Globalization.NumberFormatInfo ni = (System.Globalization.NumberFormatInfo)
            ci.NumberFormat.Clone();
            ni.NumberDecimalSeparator = ".";
            return Double.Parse(StrDouble, ni);
        }

        public static Image SetImgOpacity(Image imgPic, float imgOpac)
        {
            Bitmap bmpPic = new Bitmap(imgPic.Width, imgPic.Height);
            Graphics gfxPic = Graphics.FromImage(bmpPic);
            System.Drawing.Imaging.ColorMatrix cmxPic = new System.Drawing.Imaging.ColorMatrix();
            cmxPic.Matrix33 = imgOpac;

            System.Drawing.Imaging.ImageAttributes iaPic = new System.Drawing.Imaging.ImageAttributes();
            iaPic.SetColorMatrix(cmxPic, System.Drawing.Imaging.ColorMatrixFlag.Default, System.Drawing.Imaging.ColorAdjustType.Bitmap);
            gfxPic.DrawImage(imgPic, new Rectangle(0, 0, bmpPic.Width, bmpPic.Height), 0, 0, imgPic.Width, imgPic.Height, GraphicsUnit.Pixel, iaPic);
            gfxPic.Dispose();
            iaPic.Dispose();

            return bmpPic;
        }


        #endregion

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Save CFG
            SaveCFG svd = new SaveCFG();
            svd.RenamerInterface = RenamerInterface.SelectedIndex;
            svd.SortAfter = SortAfter.SelectedIndex;
            svd.AddFileNo = AddFileNo.SelectedIndex;
            svd.AfterRename = AfterRename.SelectedIndex;

            svd.WatermarkOverwrite = WatermarkOverwrite.SelectedIndex;
            svd.AfterWatermarked = AfterWatermarked.SelectedIndex;

            svd.ResizeType = ResizeType.SelectedIndex;
            svd.ResizeHeight = Convert.ToInt32(ResizeHeight.Text.Trim());
            svd.ResizeWidth = Convert.ToInt32(ResizeWidth.Text.Trim());
            svd.AspectRatio = AspectRatio.SelectedIndex;
            svd.ResizeQuality = Convert.ToInt32(ResizeQuality.Text.Trim());
            svd.EraseQuality = Convert.ToInt32(EraseQuality.Text.Trim());
            svd.ResizeOVerwrite = ResizeOverwrite.SelectedIndex;
            svd.AfterResize = AfterResize.SelectedIndex;
            XmlSaved<SaveCFG>.Save(XmlSaved<int>.GetCurrentDir() + @"\presets.cfg",svd);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;
            int He = Convert.ToInt32(ResizeHeight.Text.Trim());
            int Wi = Convert.ToInt32(ResizeWidth.Text.Trim());
            if ((He < 20) || (Wi < 20))
            {
                MessageBox.Show("Can't resize to image side less 20 pixels", "Resize Images", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            if (MessageBox.Show("Do you really want to resize " + listBox1.Items.Count.ToString() + " images?", "Resize Images", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            tabControl1.SelectedIndex = 6;
            listBox2.Items.Clear();
            Console("Resize Images", true);
            Console("Files to Process: " + listBox1.Items.Count.ToString(), true);
            Console("Begin", true);

            SetProgress(0, listBox1.Items.Count);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                FileDropped fd = (FileDropped)listBox1.Items[i];
                Console(" Resize '" + fd.ToString() + "' .. ", true);
                
                System.Drawing.Image im = System.Drawing.Image.FromFile(fd.fileName);
                double cropFactor = (double)im.Width/(double)im.Height;

                int nWi = Wi;
                int nHe = He;
                if (AspectRatio.SelectedIndex == 1) nWi = (int)((double)He * cropFactor);
                if (AspectRatio.SelectedIndex == 2) nHe = (int)((double)Wi / cropFactor);
                if (AspectRatio.SelectedIndex == 3)
                {
                    if(cropFactor >= 1)
                        nWi = (int)((double)He * cropFactor);
                    else
                        nHe = (int)((double)Wi / cropFactor);
                };
                if (AspectRatio.SelectedIndex == 4)
                {
                    if (cropFactor < 1)
                        nWi = (int)((double)He * cropFactor);
                    else
                        nHe = (int)((double)Wi / cropFactor);
                };
                
                bool doResize = ResizeType.SelectedIndex == 4;
                bool doCopy = ResizeType.SelectedIndex > 1;
                if ((ResizeType.SelectedIndex == 0) || (ResizeType.SelectedIndex == 2)) doResize = (nWi < im.Width) && (nHe < im.Height);
                if ((ResizeType.SelectedIndex == 1) || (ResizeType.SelectedIndex == 3)) doResize = (nWi > im.Width) && (nHe > im.Height);

                if (doResize)
                {
                    Image result = new Bitmap(nWi, nHe);
                    using (Graphics g = Graphics.FromImage((Image)result))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(im, 0, 0, nWi, nHe);
                        g.Dispose();
                    };
                    im.Dispose();
                    im = result;
                };
                                
                string subdir = System.IO.Path.GetDirectoryName(fd.fileName) + @"\Resized\";
                if (!System.IO.Directory.Exists(subdir)) System.IO.Directory.CreateDirectory(subdir);
                string savedFN = subdir + System.IO.Path.GetFileName(fd.fileName);
                if (doCopy || doResize)
                {
                    System.Drawing.Imaging.EncoderParameters pars = new System.Drawing.Imaging.EncoderParameters(1);
                    pars.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, Convert.ToInt32(ResizeQuality.Text.Trim()));
                    System.Drawing.Imaging.ImageCodecInfo ici = getEncoderInfo("image/jpeg");
                    im.Save(savedFN, ici, pars);
                    im.Dispose();
                };
                if (ResizeOverwrite.SelectedIndex == 1)
                {
                    System.IO.File.Delete(fd.fileName);
                    System.IO.File.Move(subdir + System.IO.Path.GetFileName(fd.fileName), fd.fileName);
                };
                //
                Console("OK", false);
                SetProgress(i + 1, listBox1.Items.Count);

                if (AfterResize.SelectedIndex == 1)
                    for (int x = 0; x < listBox1.Items.Count; x++)
                    {
                        FileDropped f2d = (FileDropped)listBox1.Items[x];
                        if (f2d.fileName == fd.fileName)
                        {
                            f2d.fileName = savedFN;
                            f2d.subdir += @"Resized\";
                            listBox1.Items[x] = f2d;
                        };
                    };
            };
            Console("Total Files: " + listBox1.Items.Count.ToString(), true);
            Console("End", true);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;

            if (MessageBox.Show("Do you really want to erase Exif in " + listBox1.Items.Count.ToString() + " images?", "Erase Exif", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            tabControl1.SelectedIndex = 6;
            listBox2.Items.Clear();
            Console("Erase Exif", true);
            Console("Files to Process: " + listBox1.Items.Count.ToString(), true);
            Console("Begin", true);

            SetProgress(0, listBox1.Items.Count);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                FileDropped fd = (FileDropped)listBox1.Items[i];
                Console(" Erase Exif '" + fd.ToString() + "' .. ", true);

                System.Drawing.Image im = System.Drawing.Image.FromFile(fd.fileName);
                
                System.Drawing.Imaging.EncoderParameters pars = new System.Drawing.Imaging.EncoderParameters(1);
                pars.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, Convert.ToInt32(EraseQuality.Text.Trim()));
                System.Drawing.Imaging.ImageCodecInfo ici = getEncoderInfo("image/jpeg");
                im.Save(fd.fileName+".tmp", ici, pars);
                im.Dispose();
                System.IO.File.Delete(fd.fileName);
                System.IO.File.Move(fd.fileName + ".tmp", fd.fileName);
                
                Console("OK", false);
                SetProgress(i + 1, listBox1.Items.Count);
            };
            Console("Total Files: " + listBox1.Items.Count.ToString(), true);
            Console("End", true);
        }
    }

    public struct SaveCFG
    {
        public int RenamerInterface;
        public int SortAfter;
        public int AddFileNo;
        public int AfterRename;

        public int WatermarkOverwrite;
        public int AfterWatermarked;

        public int ResizeType;
        public int ResizeHeight;
        public int ResizeWidth;
        public int AspectRatio;
        public int ResizeQuality;
        public int ResizeOVerwrite;
        public int AfterResize;

        public int EraseQuality;
    }

    /// <summary>
    ///     listBox1 obj class
    /// </summary>
    public class FileDropped
    {
        public string fileName;
        public string subdir = "";

        public FileDropped(string fileName)
        {
            this.fileName = fileName;
        }

        public override string ToString()
        {
            return System.IO.Path.GetFileName(fileName + (this.subdir.Length > 0 ? "  in  [" + this.subdir.Replace(@"\"," ").Trim() + "]" : ""));
        }
    }

    /// <summary>
    ///     ExifValue
    /// </summary>
    public class ExifValueItm
    {
        private string ekey = "";
        private string evalue = "";

        public ExifValueItm()
        {
            
        }
        
        public ExifValueItm(string Key, string Value)
        {
            this.ekey = Key;
            this.evalue = Value;
        }
        
        public string Key
        {
            get { return ekey; }
            set { ekey = value; }
        }
        
        public string Value
        {
            get { return evalue; }
            set { evalue = value; }
        }
    }

    public struct WatermarkCfg
    {
        public string Text;
        public string Image;
        public double Image2Font;
        public int ImageOpacity;
        public int ImageQuality;
        public int Opacity;
        public string FontName;
        public int FontStyle;
        public int WatermarkStyle;
    }

    /// <summary>
    ///     File Renamer Struct
    /// </summary>
    public struct fileRN : IComparer<fileRN>
    {
        public string fullPath;
        public string oldName;
        public string newName;

        public fileRN(string fullPath, string oldName, string newName)
        {
            this.fullPath = fullPath;
            this.oldName = oldName;
            this.newName = newName;
        }

        public int Compare(fileRN a, fileRN b)
        {
            return string.Compare(a.newName, b.newName);
        }
    }
}

namespace PardesiServices.WinControls
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    [ComVisible(true)]
    public class BROWSEINFO
    {
        public IntPtr hwndOwner;
        public IntPtr pidlRoot;
        public IntPtr pszDisplayName;
        public string lpszTitle;
        public int ulFlags;
        public IntPtr lpfn;
        public IntPtr lParam;
        public int iImage;
    }

    public class Win32SDK
    {
        [DllImport("shell32.dll", PreserveSig = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SHBrowseForFolder(BROWSEINFO bi);

        [DllImport("shell32.dll", PreserveSig = true, CharSet = CharSet.Auto)]
        public static extern bool SHGetPathFromIDList(IntPtr pidl, IntPtr pszPath);

        [DllImport("shell32.dll", PreserveSig = true, CharSet = CharSet.Auto)]
        public static extern int SHGetSpecialFolderLocation(IntPtr hwnd, int csidl, ref IntPtr ppidl);
    }

    [Flags, Serializable]
    public enum BrowseFlags
    {
        BIF_DEFAULT = 0x0000,
        BIF_BROWSEFORCOMPUTER = 0x1000,
        BIF_BROWSEFORPRINTER = 0x2000,
        BIF_BROWSEINCLUDEFILES = 0x4000,
        BIF_BROWSEINCLUDEURLS = 0x0080,
        BIF_DONTGOBELOWDOMAIN = 0x0002,
        BIF_EDITBOX = 0x0010,
        BIF_NEWDIALOGSTYLE = 0x0040,
        BIF_NONEWFOLDERBUTTON = 0x0200,
        BIF_RETURNFSANCESTORS = 0x0008,
        BIF_RETURNONLYFSDIRS = 0x0001,
        BIF_SHAREABLE = 0x8000,
        BIF_STATUSTEXT = 0x0004,
        BIF_UAHINT = 0x0100,
        BIF_VALIDATE = 0x0020,
        BIF_NOTRANSLATETARGETS = 0x0400,
    }

    public class FolderBrowser : Component
    {
        private string m_strDirectoryPath;
        private string m_strTitle;
        private string m_strDisplayName;
        private BrowseFlags m_Flags;
        public FolderBrowser()
        {
            m_Flags = BrowseFlags.BIF_DEFAULT;
            m_strTitle = "";
        }

        public string DirectoryPath
        {
            get { return this.m_strDirectoryPath; }
        }

        public string DisplayName
        {
            get { return this.m_strDisplayName; }
        }

        public string Title
        {
            set { this.m_strTitle = value; }
        }

        public BrowseFlags Flags
        {
            set { this.m_Flags = value; }
        }

        public DialogResult ShowDialog()
        {
            BROWSEINFO bi = new BROWSEINFO();
            bi.pszDisplayName = IntPtr.Zero;
            bi.lpfn = IntPtr.Zero;
            bi.lParam = IntPtr.Zero;
            bi.lpszTitle = "Select Folder";
            IntPtr idListPtr = IntPtr.Zero;
            IntPtr pszPath = IntPtr.Zero;
            try
            {
                if (this.m_strTitle.Length != 0)
                {
                    bi.lpszTitle = this.m_strTitle;
                }
                bi.ulFlags = (int)this.m_Flags;
                bi.pszDisplayName = Marshal.AllocHGlobal(256);
                // Call SHBrowseForFolder
                idListPtr = Win32SDK.SHBrowseForFolder(bi);
                // Check if the user cancelled out of the dialog or not.
                if (idListPtr == IntPtr.Zero)
                {
                    return DialogResult.Cancel;
                }

                // Allocate ncessary memory buffer to receive the folder path.
                pszPath = Marshal.AllocHGlobal(256);
                // Call SHGetPathFromIDList to get folder path.
                bool bRet = Win32SDK.SHGetPathFromIDList(idListPtr, pszPath);
                // Convert the returned native poiner to string.
                m_strDirectoryPath = Marshal.PtrToStringAuto(pszPath);
                this.m_strDisplayName = Marshal.PtrToStringAuto(bi.pszDisplayName);
            }
            catch
            {
                return DialogResult.Abort;
            }
            finally
            {
                // Free the memory allocated by shell.
                if (idListPtr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(idListPtr);
                }
                if (pszPath != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pszPath);
                }
                if (bi != null)
                {
                    Marshal.FreeHGlobal(bi.pszDisplayName);
                }
            }
            return DialogResult.OK;
        }

        private IntPtr GetStartLocationPath()
        {
            return IntPtr.Zero;
        }
    }
}