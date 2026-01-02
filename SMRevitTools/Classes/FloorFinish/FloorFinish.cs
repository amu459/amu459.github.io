using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using SMRevitTools.Classes.MeetCeiling;
using SMRevitTools.Classes.EnlargedDrawing;
using SMRevitTools.Classes.FloorFinish;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class FloorFinish : IExternalCommand
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

                //find the room boundary vertices for floor //Make curves and curve loop
                //Get selected rooms in Revit
                List<Room> roomList = RoomMethods.GetRoomsFromSelection(uidoc);
                Room room = roomList.FirstOrDefault();
                List<XYZ> roomVertices = RoomMethods.GetRoomVertex(room);



                //Get floor types in Revit
                //Display floor types in a text box and ask user for select the floor finish to be modelled
                //Collect all floors types from the projects
                List<ElementType> floorTypes = collector.WhereElementIsElementType().OfCategory(BuiltInCategory.OST_Floors).Cast<ElementType>().ToList();

                //List all the floor type names from the project
                List<string> floorNames = floorTypes.Select(fl => fl.Name).ToList();

                string[] floorNamesString = floorNames.ToArray();
                //CollectUserData

                FloorSelection inputWindow = new FloorSelection(uidoc, floorNamesString);
                inputWindow.ShowDialog();

                string floorTypeNameSelected = inputWindow.inputFloorType;
                string offsetGiven = inputWindow.inputOffset;
                double offsetParsed;
                bool isParsed = Double.TryParse(offsetGiven, out offsetParsed);
                if(!isParsed)
                {
                    TaskDialog.Show("Revit", "Incoorect offset value provided??"
                    +Environment.NewLine + offsetGiven + " - What is this??") ;
                    goto skipped;
                }

                offsetParsed = offsetParsed / 304.8;

                //TaskDialog.Show("Revit", "Floor Type selecterd = \"" + floorTypeNameSelected + "\" and offset selected = " + offsetGiven);


                //Find floor finish type
                ElementType floorTypeSelected = floorTypes.Where(x => x.Name == floorTypeNameSelected).FirstOrDefault();



                // Create rooms to specific family locations
                using (Transaction transaction = new Transaction(doc, "Floor by Room"))
                {
                    transaction.Start();
                    CurveLoop roomVertexLoop = new CurveLoop();
                    int totalVertices = roomVertices.Count;

                    roomVertices.Add(roomVertices.First());
                    for (int i = 0; i < totalVertices; i++) 
                    {
                        roomVertexLoop.Append(Line.CreateBound(roomVertices[i], roomVertices[i + 1]));
                    }

                    Floor floorModelled = Floor.Create(doc, new List<CurveLoop> { roomVertexLoop }, floorTypeSelected.Id, room.Level.Id);
                    Parameter paramEle = floorModelled.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    paramEle.Set(offsetParsed);
                    transaction.Commit();
                }

                // Return success result

                string toolName = "Floor by room";
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
                string toolName = "Floor by room";
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
