using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace ToolsV2Classes.Class.Lockers
{
    /// <summary>
    /// Interaction logic for LockerOutput.xaml
    /// </summary>
    public partial class LockerOutput : Window
    {
        public UIDocument uidoc { get; }
        public Document doc { get; set; }

        public LockerOutput()
        {
            InitializeComponent();
        }
    }
}
