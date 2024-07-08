using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;


namespace ToolsV2Classes.Class.FFnE
{
    /// <summary>
    /// Interaction logic for packBInputWindow.xaml
    /// </summary>
    public partial class PlantInputWindow : Window
    {
        public UIDocument uidoc { get; }
        public Document doc { get; set; }
        public string inputText { get; set; }
        public string inputLevelName { get; set; }



        public PlantInputWindow(UIDocument uiDoc, string[] levelNames)
        {
            uidoc = uiDoc;
            doc = uiDoc.Document;
            InitializeComponent();
            string[] listLevelsTemp = new string[2] { "level 1", "level 2" };
            exCombo.ItemsSource = levelNames;
            Title = "Input Casework Name";
        }

        private void GetPlantName(object sender, RoutedEventArgs e)
        {
            inputText = inputCaseworkText.Text;
            if (exCombo.SelectedItem != null)
            {
                inputLevelName = exCombo.SelectedItem.ToString();
            }
            else
            {
                inputLevelName = "00";
            }
            Close();
            return;
        }

        private void CancelTask(object sender, RoutedEventArgs e)
        {
            inputText = "cancel";
            inputLevelName = "00";
            Close();
            return;
        }

        private void inputCaseworkText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
