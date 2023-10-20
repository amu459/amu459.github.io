#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion // Namespaces

namespace ToolsV2Classes
{

    // 3D Rooms Deletion
    [Transaction(TransactionMode.Manual)]
    public class RoomMassDelete : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view;
            view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            // Deleting existing DirectShape
            // get ready to filter across just the elements visible in a view 
            FilteredElementCollector coll = new FilteredElementCollector(doc, view.Id);
            coll.OfClass(typeof(DirectShape));
            IEnumerable<DirectShape> DSdelete = coll.Cast<DirectShape>();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Delete elements");
                try
                {
                    foreach (DirectShape ds in DSdelete)
                    {
                        ICollection<ElementId> ids = doc.Delete(ds.Id);
                    }
                    tx.Commit();
                }
                catch (ArgumentException)
                {
                    tx.RollBack();
                }
            }

            string toolName = "3DrOOms Destroy";
            DateTime endTime = DateTime.Now;
            var deltaTime = endTime - startTime;
            var detlaMilliSec = deltaTime.Milliseconds;
            UIApplication uiApp = commandData.Application;
            HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
            return Result.Succeeded;
        }
    }

}