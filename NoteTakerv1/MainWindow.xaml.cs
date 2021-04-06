using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Words.NET;
using GI.Screenshot;
using Xceed.Document.NET;
using System.IO;

namespace NoteTakerv1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //private string textboxip = "";
        //private DocX doc;
        //private string filepath = "";
        public MainWindow()
        {
            InitializeComponent();
            //first change; checking version control.

        }

        private void ScreenShotButton_Click(object sender, RoutedEventArgs e)
        {

            //filepath = FileOp.getFileName();

            if (FileOp.isNameEmpty())
            {
                RemarksLabel.Content = "please select a file first";
            }
            else
            {
                BitmapSource image1 = Screenshot.CaptureAllScreens();
                //BitmapFrame bf = BitmapFrame.Create(image);

                //Stream logo = new MemoryStream();
                FileOp.addPicture(image1);

                RemarksLabel.Content = "screenshotted";
                //MainTextBox.Clear();
            }

        }

        private void SnipButton_Click(object sender, RoutedEventArgs e)
        {

            if (FileOp.isNameEmpty())
            {
                RemarksLabel.Content = "please select a file first";
            }
            else
            {
                //textboxip = MainTextBox.Text;
                BitmapSource image1 = Screenshot.CaptureRegion();

                FileOp.addPicture(image1);

                RemarksLabel.Content = "snipping";
                //MainTextBox.Clear();
            }

        }

        private void PushTextButton_Click(object sender, RoutedEventArgs e)
        {
            string textboxip;

            if(FileOp.isNameEmpty())
            {
                RemarksLabel.Content = "please select a file first";
            }
            else
            {
                textboxip = MainTextBox.Text;
                //doc = DocX.Load(filepath);
                FileOp.addText(textboxip);
                RemarksLabel.Content = "pushed";
                MainTextBox.Clear();
            }
            

        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            
            //i'm using only word documents, so..
            openFileDialog.Filter = "Word Documents|*.docx";
            
            //did the dialog open up; stored in response; nullable bool
            bool? response = openFileDialog.ShowDialog();

            if(response == true)
            {
                //filename stored here, in filepath
                string filepath = openFileDialog.FileName;
                FileOp.setFileName(filepath);

                RemarksLabel.Content = "selected file is" + filepath;

                FileOp.openFile();

            }

            if(FileOp.isNameEmpty()) //empty filepath
            {
                RemarksLabel.Content = "please select a file";
            } 

        }

        private void LaunchFileButton_Click(object sender, RoutedEventArgs e)
        {

            if (FileOp.isNameEmpty())
            {
                RemarksLabel.Content = "please select a file first";
            }
            else
            {
                string filepath = FileOp.getFileName();
                //RemarksLabel.Content = filepath;
                //Console.WriteLine(filepath);
                //string filename = @"C:\Users\Gaurav\Desktop\forNoteTaker.docx";

                string op = FileOp.launchFile();

                RemarksLabel.Content = op;
                //MainTextBox.Clear();
            }

        }
    }
}
