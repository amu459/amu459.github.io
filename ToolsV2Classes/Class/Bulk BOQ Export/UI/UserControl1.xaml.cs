using Autodesk.Revit.UI;
using System.Windows;

namespace ToolsV2Classes
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : System.Windows.Controls.UserControl
    {

        Autodesk.Revit.DB.Document doc { get; set; }
        UIApplication application { get; set; }
        public UserControl1(Autodesk.Revit.DB.Document document,UIApplication uIApplication)
        {
            doc = document;
            application= uIApplication;
            InitializeComponent();
        }
        UIOperation uIOperation = new UIOperation();
        private void CalculateBOQ_Click(object sender, RoutedEventArgs e)
        {

            uIOperation.calculateBOQ(doc, ExcelPath.Text, application);
            (this.Parent as Window).Close();
        }
        private void BrowsePath_Click(object sender, RoutedEventArgs e)
        {
            //UIOperation uIOperation = new UIOperation();
            uIOperation.browsePath(this);
        }
    }
}
