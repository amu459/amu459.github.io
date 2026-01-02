using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class ClassTemplate : IExternalCommand
    {

        // Implement the Execute method
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            try
            {
                // Create a filtered element collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);






                // Create rooms to specific family locations
                using (Transaction transaction = new Transaction(doc, "Tool Name"))
                {
                    transaction.Start();


                    transaction.Commit();
                }

                // Return success result

                string toolName = "Tool Name";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success - ", doc, uiApp, detlaMilliSec);
            skipped:
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Tool Name";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }
    }
}
