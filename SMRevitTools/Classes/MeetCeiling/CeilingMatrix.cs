

/*
 * Grid ceiling + Gypsum ceiling floating
 * Innerloop vertices will be used in this class for floating gypsum ceiling
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using SMRevitTools.Classes.MeetCeiling;
using System.Windows.Documents;
using Autodesk.Revit.Creation;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CeilingMatrix : IExternalCommand
    {

        // Implement the Execute method
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            try
            {
                // Create a filtered element collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                //Get selected rooms in Revit
                List<Room> roomList = RoomMethods.GetRoomsFromSelection(uidoc);
                Room room = roomList.FirstOrDefault();
                List<XYZ> roomVertices = RoomMethods.GetRoomVertex(room);
                // Step 1: Find bottom-left point (smallest y, then smallest x)
                XYZ bottomLeft = roomVertices.OrderBy(p => p.Y).ThenBy(p => p.X).First();

                // Step 2: Sort remaining points based on angle from bottom-left (clockwise)
                List<XYZ> sortedPoints = roomVertices.Where(p => p != bottomLeft).OrderBy(p => Math.Atan2(p.Y - bottomLeft.Y, p.X - bottomLeft.X)).ToList();
                // Step 3: Ensure bottom-left is the first element
                sortedPoints.Insert(0, bottomLeft);
                roomVertices = sortedPoints;
                var realRoomVertices = sortedPoints;

                //Bounding Box of Room Vertices
                //roomVertices = RoomMethods.GetBoundingBox(roomVertices);
                
                //Angle of Room Rotation
                double rotationAngle = RoomMethods.FindRotationAngle(roomVertices);
                if (rotationAngle != 0 || rotationAngle != Math.PI * 0.5 || rotationAngle != Math.PI)
                {
                    //Transform
                    Transform tForm = Transform.CreateRotation(XYZ.BasisZ, -rotationAngle);
                    roomVertices = RoomMethods.TransformList(tForm, roomVertices);
                    roomVertices = RoomMethods.GetBoundingBox(roomVertices);

                }
                else
                {
                    roomVertices = RoomMethods.GetBoundingBox(roomVertices);
                }

                List <XYZ> gridRoomVertices = RoomMethods.GetGridRoomVertex(roomVertices);
                List<XYZ> gridInnerVertices = RoomMethods.GetInnerOffsetVertex(gridRoomVertices);




                #region Checking with Prompt
                //string prompt = "Room vertices = " + Environment.NewLine;
                //foreach (XYZ pt in roomVertices)
                //{
                //    prompt += "X = " + Math.Round(pt.X, 2).ToString() + " & Y = " + Math.Round(pt.Y, 2).ToString() + Environment.NewLine;

                //}

                //prompt += Environment.NewLine + "Offset Vertices = " + Environment.NewLine;
                //foreach (XYZ pt in gridRoomVertices)
                //{
                //    prompt += "X = " + Math.Round(pt.X, 2).ToString() + " & Y = " + Math.Round(pt.Y, 2).ToString() + Environment.NewLine;
                //}

                //TaskDialog.Show("Revit", prompt);
                #endregion

                if (rotationAngle != 0 || rotationAngle != Math.PI * 0.5 || rotationAngle != Math.PI)
                {
                    //Transform back
                    Transform tFormBack = Transform.CreateRotation(XYZ.BasisZ, rotationAngle);
                    roomVertices = RoomMethods.TransformList(tFormBack, roomVertices);
                    gridRoomVertices = RoomMethods.TransformList(tFormBack, gridRoomVertices);
                    gridInnerVertices = RoomMethods.TransformList(tFormBack, gridInnerVertices);

                }

                //TaskDialog.Show("Revit", "Angle = " + (rotationAngle * 180/Math.PI).ToString());




                roomVertices.Add(roomVertices.First());
                gridRoomVertices.Add(gridRoomVertices.First());
                gridInnerVertices.Add(gridInnerVertices.First());

                ElementType gypsumCeilingType = collector.WhereElementIsElementType().OfCategory(BuiltInCategory.OST_Ceilings).Cast<ElementType>().Where(x => x.Name.ToLower().Contains("gypsum")).FirstOrDefault();

                ElementType gridCeilingType = collector.WhereElementIsElementType().OfCategory(BuiltInCategory.OST_Ceilings).Cast<ElementType>().Where(x => x.Name.ToLower().Contains("600x600")).FirstOrDefault();

                ElementId gypsumCeilingId = gypsumCeilingType.Id;
                ElementId gridCeilingId = gridCeilingType.Id;




                
                using (Transaction transaction = new Transaction(doc, "Ceiling Matrix"))
                {
                    transaction.Start();
                    CurveLoop roomVertexLoop = new CurveLoop();
                    CurveLoop gridRoomVertexLoop = new CurveLoop();
                    CurveLoop gridInnerVertexLoop = new CurveLoop();
                    for (int i = 0; i < 4; i++)
                    {
                        roomVertexLoop.Append(Line.CreateBound(roomVertices[i], roomVertices[i + 1]));
                        gridRoomVertexLoop.Append(Line.CreateBound(gridRoomVertices[i], gridRoomVertices[i + 1]));
                        gridInnerVertexLoop.Append(Line.CreateBound(gridInnerVertices[i], gridInnerVertices[i + 1]));
                    }


                    var gypCeiling = Ceiling.Create(doc, new List<CurveLoop> { roomVertexLoop, gridInnerVertexLoop}, gypsumCeilingId, room.Level.Id);
                    Parameter paramEle = gypCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                    paramEle.Set(2700/304.8);

                    var gridCeiling = Ceiling.Create(doc, new List<CurveLoop> { gridRoomVertexLoop }, gridCeilingId, room.Level.Id);
                    Parameter paramEle2 = gridCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                    paramEle2.Set(2550 / 304.8);

                    transaction.Commit();
                }

                // Return success result

                string toolName = "Ceiling Matrix";
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
                string toolName = "Ceiling Matrix";
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
