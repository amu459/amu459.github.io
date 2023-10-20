using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows;

namespace ToolsV2Classes
{

    [Transaction(TransactionMode.Manual)]
    public class SyncAirTableQuickClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;
            Window mainWindow = new Window
            {
                Title = "Calculate BOQ",
                Content = new UserControl1(doc,uiApp)
           
            };
            mainWindow.Height = 250;
            mainWindow.Width = 550;
            mainWindow.MaxHeight = 250;
            mainWindow.MinHeight = 250;
            mainWindow.MaxWidth = 525;
            mainWindow.MinWidth = 525;
            mainWindow.ShowDialog();


           

            return Result.Succeeded;
        }

    }

}
