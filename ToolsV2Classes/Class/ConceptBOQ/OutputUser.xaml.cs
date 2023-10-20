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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ToolsV2Classes.Class.ConceptBOQ
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>

    public partial class OutputUser : Window
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        public string boqFilePath = "";
        public OutputUser(ExternalCommandData commandData, string boqPath)
        {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
            boqFilePath = boqPath;
        }

        private void LoadingBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            LoadingBox.Text = boqFilePath;
        }
    }
}
