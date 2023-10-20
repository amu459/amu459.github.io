using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;


namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RoomSeparation : IExternalCommand
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
                // Create a filtered element collector to get all room separation lines
                FilteredElementCollector roomSeparationLinesCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_RoomSeparationLines);

                List<ElementId> allRSIds = new List<ElementId>();
                foreach(Element rS in roomSeparationLinesCollector)
                {
                    allRSIds.Add(rS.Id);
                }


                // Create a FilteredElementCollector to collect all rooms in the project
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Rooms);

                // Create a list to store the rooms
                List<Room> rooms = new List<Room>();

                // Loop through the collected rooms and add them to the list
                foreach (Element element in collector)
                {
                    Room room = element as Room;
                    if (room != null && room.Area > 1)
                    {
                        rooms.Add(room);
                    }
                }

                List<ElementId> roomSeprationIds = new List<ElementId>();

                foreach (Room r in rooms)
                {
                    IList<IList<BoundarySegment>> loops = DeskAutomation.Helper_Classes.RoomData.GetBoundaryLoops(r);
                    foreach(IList<BoundarySegment> loop in loops)
                    {
                        foreach(BoundarySegment seg in loop)
                        {
                            ElementId segId = seg.ElementId;
                            Element segEle = doc.GetElement(segId);

                            if(segEle != null)
                            {
                                Category segCat = segEle.Category;
                                string catName = segCat.Name;
                                
                                if (catName.ToLower().Contains("separation"))
                                {
                                    roomSeprationIds.Add(segId);
                                }
                            }

                        }
                    }
                }

                List<ElementId> uniqueRSIds = roomSeprationIds.Distinct().ToList();
                List<ElementId> unwantedRSIds = allRSIds.Except(uniqueRSIds).ToList();

                TaskDialog.Show("Revit Window:", "# Deleted room separation lines = " + unwantedRSIds.Count().ToString()
                    + Environment.NewLine + "# total lines before deletion = " + allRSIds.Count().ToString()
                    + Environment.NewLine + "# useful lines = " + uniqueRSIds.Count().ToString());


                // Delete each unnecessary room separation line
                using (Transaction transaction = new Transaction(doc, "Delete Room Separation"))
                {
                    transaction.Start();
                    foreach (ElementId elId in unwantedRSIds)
                    {

                        if (doc.GetElement(elId).Pinned)
                        {
                            doc.GetElement(elId).Pinned = false;
                        }
                        doc.Delete(elId);
                    }

                    transaction.Commit();
                }

                // Return success result

                string toolName = "Clean RS";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Clean RS";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);

                message = e.Message;
                return Result.Failed;
            }
            
        }
    }
}