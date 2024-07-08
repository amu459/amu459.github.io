using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;


namespace BOQStandardCheck
{
    /// <summary>
    /// Interaction logic for BOQStandardCheckUI.xaml
    /// </summary>
    public partial class BOQStandardCheckUI : UserControl
    {
        ErrorData errorData { get; set; }
        public BOQStandardCheckUI()
        {
            InitializeComponent();
        }

        private void BrowseFile1_Click(object sender, RoutedEventArgs e)
        {
            ExcelPath1.Text = StandardBOQCheckUiOperation.browseFile();
        }

        private void BrowseFile2_Click(object sender, RoutedEventArgs e)
        {
            ExcelPath2.Text = StandardBOQCheckUiOperation.browseFile();
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            errorData = BOQStandardCheckUtilityClass.showReport(ExcelPath1.Text, ExcelPath2.Text);
            // (this.Parent as Window).Close();

        }

        private void Download_Click(object sender, RoutedEventArgs e)
        {
            errorData = BOQStandardCheckUtilityClass.compareExcel(ExcelPath1.Text, ExcelPath2.Text);
            //(this.Parent as Window).Close();
            Window mainWindow = new Window
            {
                Title = "Download BOQ Standard Check Report",
                Content = new BOQStandardCheckReportUI(errorData)
            };
            mainWindow.Height = 225;
            mainWindow.Width = 625;
            mainWindow.MaxHeight = 225;
            mainWindow.MinHeight = 225;
            mainWindow.MaxWidth = 625;
            mainWindow.MinWidth = 625;
            mainWindow.ShowDialog();

            //(this.Parent as Window).Close();


        }


    }
}
