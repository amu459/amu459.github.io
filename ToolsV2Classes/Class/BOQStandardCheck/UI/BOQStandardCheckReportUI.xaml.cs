using System;
using System.Collections.Generic;
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

namespace BOQStandardCheck
{
    /// <summary>
    /// Interaction logic for BOQStandardCheckReportUI.xaml
    /// </summary>
    public partial class BOQStandardCheckReportUI : UserControl
    {
        ErrorData errData { get; set; }
        public BOQStandardCheckReportUI(ErrorData errorData)
        {
            errData = errorData;
            InitializeComponent();
        }

        private void DownloadReport_Click(object sender, RoutedEventArgs e)
        {
            BOQStandardCheckUtilityClass.writeCompareReport(errData, ReportExcelPath.Text);
            (this.Parent as Window).Close();
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            ReportExcelPath.Text = StandardBOQCheckUiOperation.browsePath();
            //(this.Parent as Window).Close();
        }
    }
}
