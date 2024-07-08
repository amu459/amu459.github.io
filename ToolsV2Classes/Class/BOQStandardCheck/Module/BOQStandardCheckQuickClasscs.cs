using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;

namespace BOQStandardCheck
{

    [Transaction(TransactionMode.Manual)]
    public class BOQStandardCheckQuickClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Window mainWindow = new Window
            {
                Title = "Compare BOQ Standard", 
                Content = new BOQStandardCheckUI()
            };
            mainWindow.Height = 325;
            mainWindow.Width = 625;
            mainWindow.MaxHeight = 325;
            mainWindow.MinHeight = 325;
            mainWindow.MaxWidth = 625;
            mainWindow.MinWidth = 625;
            mainWindow.ShowDialog();
            return Result.Succeeded;
        }
    }
}
