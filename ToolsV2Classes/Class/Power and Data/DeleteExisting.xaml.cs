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

namespace ToolsV2Classes.Class.Power_and_Data
{
    /// <summary>
    /// Interaction logic for DeleteExisting.xaml
    /// </summary>
    public partial class DeleteExisting : Window
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        public bool grommetDelete;
        public bool canceled = false;
        public DeleteExisting(ExternalCommandData commandData)
        {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            grommetDelete = true;
            Close();
            return;

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            grommetDelete = false;
            Close();
            return;
        }
    }
}
