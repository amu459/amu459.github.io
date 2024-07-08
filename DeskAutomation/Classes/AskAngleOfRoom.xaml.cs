using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace DeskAutomation.Classes
{
    /// <summary>
    /// Interaction logic for AskAngleOfRoom.xaml
    /// </summary>
    public partial class AskAngleOfRoom : Window
    {
        public string inputAngle { get; set; }


        public AskAngleOfRoom()
        {
            InitializeComponent();
        }


        private void CancelTask(object sender, RoutedEventArgs e)
        {
            inputAngle = "cancel";
            Close();
            return;
        }

        private void GetAngleValue(object sender, RoutedEventArgs e)
        {
            inputAngle = inputAngleText.Text;

            Close();
            return;
        }
    }
}
