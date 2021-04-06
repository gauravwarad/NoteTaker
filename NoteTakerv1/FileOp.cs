using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Media.Imaging;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace NoteTakerv1
{
    static class FileOp
    {
        private static string filename = "";
        private static DocX doc;

        public static void openFile()
        {
            doc = DocX.Load(filename);
            doc.InsertParagraph("");
            doc.Save();
        }

        public static void addText(string textboxip)
        {
            doc = DocX.Load(filename);
            doc.InsertParagraph(textboxip);
            doc.Save();
        }

        public static void addPicture(BitmapSource image1)
        {
            string username = Environment.UserName;
            string imgPath = @"C:\Users\"+ username + @"\Desktop\banana.png";
            using (var fileStream = new FileStream(imgPath, FileMode.Create))
            {
                BitmapEncoder encoder = new PngBitmapEncoder();
                //encoder.Frames.Add(BitmapFrame.Create(image));
                encoder.Frames.Add(BitmapFrame.Create(image1));
                encoder.Save(fileStream);
            }


            var doc = DocX.Load(filename);
            Xceed.Document.NET.Image img = doc.AddImage(imgPath);
            Picture p = img.CreatePicture();
            Xceed.Document.NET.Paragraph para = doc.InsertParagraph("");
            para.AppendPicture(p);

            doc.Save();
            File.Delete(imgPath);
        }
        public static string launchFile()
        {
            try
            {
                Process.Start(@"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE", filename);
                return "launching";
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
        }
        public static string getFileName()
        {
            return filename;
        }
        
        public static void setFileName(string filename)
        {
            FileOp.filename = filename;
        }
        public static bool isNameEmpty()
        {
            if (filename == "")
            
                return true;
            
            else
                return false;
        }
    }
}

//application developed by Gaurav Warad
//find me at github.com/gauravwarad or @GauravWarad on twitter
//enjoy.