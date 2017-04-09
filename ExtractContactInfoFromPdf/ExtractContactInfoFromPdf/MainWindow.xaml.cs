using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace ExtractContactInfoFromPdf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<ContactInfo> contactInfoList = new List<ContactInfo>();
        public MainWindow()
        {
            InitializeComponent();
            //validateExpiration();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            int textBoxRightMargin = 30, textBoxBottonMargin = 30 + 23;
            System.Windows.Point windowPosition = PointToScreen(new System.Windows.Point(0, 0));

            System.Windows.Point textBoxPosition = textBox.PointToScreen(new System.Windows.Point(0, 0));
            textBox.Width = e.NewSize.Width - (textBoxPosition.X - windowPosition.X) - textBoxRightMargin;
            textBox.Height = e.NewSize.Height - (textBoxPosition.Y - windowPosition.Y) - textBoxBottonMargin;
        }

        private void openButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf) | *.pdf";
            openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                openButton.IsEnabled = false;
                string titleBackup = Title;
                extractData(openFileDialog.FileName);
                Title = titleBackup;
                openButton.IsEnabled = true;
            }
        }

        void validateExpiration()
        {
            DateTime expireDate = new DateTime(2017, 3, 20);
            DateTime today = DateTime.Now;
            if (today > expireDate)
            {
                Application.Current.Shutdown();
            }
        }

        void extractData(string pdfFilePath)
        {
            textBox.Clear();
            float /*pageWidth = 612, */pageHeight = 792;
            float width = 250, height = 140;
            float[] xCoords = { 35, 300 };
            float[] yCoords = { 200, 370, 540, 710 };
            try
            {
                int startPageNumber = int.Parse(Regex.Replace(startPageNumberTextBox.Text, @"[^0-9]+", ""));
                int pageNumbers = Math.Min(int.Parse(Regex.Replace(pageNumbersTextBox.Text, @"[^0-9]+", "")), 100);
                int endPageNumber;
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    //var pageRectangle = reader.GetPageSize(1);
                    //MessageBox.Show(string.Format("{0}:{1}", pageRectangle.Width, pageRectangle.Height));
                    endPageNumber = Math.Min(reader.NumberOfPages, startPageNumber + pageNumbers - 1);
                    for (int pageNo = startPageNumber; pageNo <= endPageNumber; pageNo++)
                    {
                        double percentageOfProgress = (double)(pageNo - startPageNumber + 1) * 100.0 / (endPageNumber - startPageNumber + 1);
                        Title = string.Format("[{0}%] Processing ... [CurrentPageNo: {1}]", percentageOfProgress.ToString(".0"), pageNo);
                        for (int i = 0; i < xCoords.Length; i++)
                        {
                            for (int j = 0; j < yCoords.Length; j++)
                            {
                                float x = xCoords[i];
                                float y = pageHeight - yCoords[j];
                                RenderFilter[] filter = { new RegionTextRenderFilter(new System.util.RectangleJ(x, y, width, height)) };
                                ITextExtractionStrategy strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);
                                string plainText = PdfTextExtractor.GetTextFromPage(reader, pageNo, strategy);
                                ContactInfo contactInfo = new ContactInfo();
                                if (contactInfo.tryGetContactInfo(plainText))
                                {
                                    contactInfoList.Add(contactInfo);
                                }
                            }
                        }
                    }
                }
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = "CSV files (*.csv) | *.csv";
                saveFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
                saveFileDialog.FileName = string.Format("{0}_{1}.csv", startPageNumber.ToString("000"), endPageNumber.ToString("000"));
                if (saveFileDialog.ShowDialog() == true)
                {
                    StreamWriter sw = new StreamWriter(saveFileDialog.FileName);
                    for (int i = 0; i < contactInfoList.Count; i++)
                    {
                        string contactInfoString = contactInfoList[i].toString();
                        sw.Write(contactInfoString);
                        if (i <= 100) textBox.Text += contactInfoString;
                    }
                    sw.Close();
                }
            }
            catch { }
            MessageBox.Show("Success");
        }
    }

    class ContactInfo
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string CellNumber { get; set; }
        public string Email { get; set; }

        public ContactInfo()
        {
            FirstName = LastName = CellNumber = Email = "";
        }

        public bool tryGetContactInfo(string plainText)
        {
            char[] lineSeparatingChars = { '\n' };
            string[] lines = plainText.Split(lineSeparatingChars, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return false;
            FirstName = lines[0];
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.StartsWith("Email:"))
                {
                    Email = line.Length > 6 ? line.Substring(7).Trim() : lines[i - 1].Trim();
                }
                if (line.StartsWith("Cell Ph:"))
                {
                    CellNumber = line.Length > 8 ? line.Substring(9).Trim() : lines[i - 1].Trim();
                }
            }
            return CellNumber.Length > 0 || Email.Length > 0;
        }

        public string toString()
        {
            return string.Format("{0},{1},{2}\n", FirstName.Replace(",", ""), CellNumber.Replace(",", ""), Email.Replace(",", ""));
        }
    }
}
