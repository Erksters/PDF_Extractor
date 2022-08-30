using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using System.Windows.Forms;
namespace PDF_Reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String PDF_Name = null;
        String PDF_SafeFileName = null;
        String Output_Path = null;
        Dictionary<String, String> UserInputDictionary = new Dictionary<String, String>();
        Boolean User_Selected_PDF = false;
        Boolean User_Selected_Output_Directory = false;
        PdfArray annotArray;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// This will open the file explorer and helps select the pdf in question
        /// Also sets up the global objects such as pdf_SafeFileName and path to pdf (PDF_Name)_
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void File_Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            openFileDlg.Filter = "PDF Documents (.pdf)|*.pdf";

            if (openFileDlg.ShowDialog() == true)
            {
                Title_Label.Content = openFileDlg.SafeFileName;
                PDF_SafeFileName = openFileDlg.SafeFileName;
                PDF_Name = openFileDlg.FileName;
                User_Selected_PDF = true;
            }

            EnableConversionButtons();
        }

        /// <summary>
        /// This will open the file explorer and helps the user select where they want to dump the json and csv files
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Folder_Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = fbd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Output_Path = fbd.SelectedPath;
                    Output_Directory_TextBlock.Text = fbd.SelectedPath + "\\";
                    User_Selected_Output_Directory = true;
                }
            }
            EnableConversionButtons();
        }

        /// <summary>
        /// Checks if the PDF has been selected and if the Output directory has been set
        /// Enables the Conversion buttons if both are good to go
        /// </summary>
        private void EnableConversionButtons()
        {
            if (User_Selected_PDF && User_Selected_Output_Directory)
            {
                Convert_Button.IsEnabled = true;
                JSON_Button.IsEnabled = true;
            }
        }

        /// <summary>
        /// Crawls pdf and extracts data. Adds that data to the global dictionary (UserInputDictionary) for later conversions
        /// Used in CSV and JSON button clicks
        /// </summary>
        private void PrepConversionCreateDictionary()
        {
            PdfReader reader = new PdfReader(PDF_Name);
            int Pages = reader.NumberOfPages;
            for (int i = 1; i < Pages +1; i++)
            {
                //Find Page annotations (if any)
                PdfDictionary pageDict = reader.GetPageN(i);
                annotArray = pageDict.GetAsArray(PdfName.ANNOTS);

                if (annotArray != null && annotArray.Length > 0)
                {
                    //Go through annotations within page
                    for (int j = 0; j < annotArray.Size ; ++j)
                    {
                        try
                        {
                            PdfDictionary curAnnot = annotArray.GetAsDict(j);
                            string AnnotationName = curAnnot.GetAsString(PdfName.T).ToString();
                            string AnnotationContents = curAnnot.GetAsString(PdfName.V) == null ? PdfString.NOTHING : curAnnot.GetAsString(PdfName.V).ToString();

                            UserInputDictionary.Add(AnnotationName.ToString(), AnnotationContents);
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("I had trouble accessing an entry, but I'll try my best");
                        }
                        

                    } // Exit list of annotations on page crawl
                }
            } // Exit PDF crawl 
        }

        /// <summary>
        /// Will convert the prepped dictionary items into a csv file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Convert_Button_Click(object sender, RoutedEventArgs e)
        {
            PrepConversionCreateDictionary();

            String csv = "Annotation_Name,Annotation_Entry\n" + String.Join(
                                        Environment.NewLine,
                                        UserInputDictionary.Select(d => $"{d.Key},{d.Value}")
                                    );

            File.WriteAllText(Output_Path+"\\"+ PDF_SafeFileName.Substring(0,PDF_SafeFileName.Length - 4) + ".csv", csv);

            System.Windows.MessageBox.Show("Created CSV file Successfully. " + "I copied " + UserInputDictionary.Count.ToString() + " Lines" + "I'm shutting down now!");
            App.Current.Shutdown();
        }

        /// <summary>
        /// Will convert the prepped dictionary items into a JSON file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JSON_Button_Click(object sender, RoutedEventArgs e)
        {
            PrepConversionCreateDictionary();
            string json = JsonConvert.SerializeObject(UserInputDictionary, Formatting.Indented);

            File.WriteAllText(Output_Path + "\\" + PDF_SafeFileName.Substring(0, PDF_SafeFileName.Length - 4) + ".json", json);

            System.Windows.MessageBox.Show("Created JSON file Successfully. " + "I copied " + UserInputDictionary.Count.ToString() + " Lines" + "I'm shutting down now!");
            App.Current.Shutdown();


        }
    }
}
