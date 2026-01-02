
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;


namespace SMRevitTools.Classes.FloorFinish
{
    /// <summary>
    /// Interaction logic for packBInputWindow.xaml
    /// </summary>
    public partial class FloorSelection : Window
    {
        public UIDocument uidoc { get; }
        public Document doc { get; set; }
        public string inputOffset { get; set; }
        public string inputFloorType { get; set; }



        public FloorSelection(UIDocument uiDoc, string[] levelNames)
        {
            uidoc = uiDoc;
            doc = uiDoc.Document;
            InitializeComponent();
            exCombo.ItemsSource = levelNames;
            Title = "Input Floor Type";
        }

        private void GetPackBName(object sender, RoutedEventArgs e)
        {
            inputOffset = inputCaseworkText.Text;
            if (exCombo.SelectedItem != null)
            {
                inputFloorType = exCombo.SelectedItem.ToString();
            }
            else
            {
                inputFloorType = "NA";
            }
            Close();
            return;
        }

        private void CancelTask(object sender, RoutedEventArgs e)
        {
            inputOffset = "cancel";
            inputFloorType = "NA";
            Close();
            return;
        }

        private void inputCaseworkText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void exCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}

