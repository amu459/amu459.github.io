﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class HorizontalLights1200 : IExternalCommand
    {
        List<FamilyInstance> lightsModeled = new List<FamilyInstance>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            double roomTotalArea = 0;

            try
            {
                #region Get Lighting Fixture Family
                //Get Lighting fixture Type for Office Space
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                FamilySymbol light1200Symbol = collector.OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(x => x.FamilyName.Contains("IN-Architecture Linear-12"))
                    .FirstOrDefault(x => x.Name.Contains("WWI-LT-AL-12-01"));

                if (light1200Symbol == null)
                {
                    TaskDialog.Show("Revit", "Standard Lighting fixture family is not loaded into the project."
                        + Environment.NewLine
                        + "Tool will try to load the latest lighting fixture Family for linear lights: IN-Architecture Linear-12");

                    using (Transaction tx = new Transaction(doc, "Load Light Fixture Family"))
                    {
                        tx.Start();
                        if (light1200Symbol == null)
                        {
                            string path = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Light Fixture\\IN-Architecture Linear-12.rfa";
                            FamilyLoadOption newOption = new FamilyLoadOption();
                            doc.LoadFamily(path, newOption, out Family lightFamily);
                            if (lightFamily == null)
                            {
                                TaskDialog.Show("Revit", "Light Family cannot be Loaded, please load the family 'IN-Architecture Linear-12' manually :(");
                            }
                            else
                            {
                                TaskDialog.Show("Revit", "Light Family Loaded:" + lightFamily.Name + Environment.NewLine + "Please run the tool again :)");
                            }
                        }
                        tx.Commit();
                    }
                    goto cleanup;
                }
                #endregion

                #region Get Rooms from Selection
                List<Room> roomList = new List<Room>();
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (0 == selectedIds.Count)
                {
                    // If no elements selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + " You haven't selected any Rooms!");
                }
                else
                {
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = uidoc.Document.GetElement(id);
                        if (elem is Room)
                        {
                            Room testRedundant = elem as Room;
                            string programType = LightingToolMethods.GetParamVal(testRedundant, "WW-ProgramType");
                            if (testRedundant.Area != 0 && programType.ToLower() == "work")
                            {
                                //Add rooms to roomList
                                roomList.Add(elem as Room);
                            }
                        }
                    }
                    if (0 == roomList.Count)
                    {
                        // If no rooms selected.
                        TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Rooms!");
                    }
                }
                #endregion

                if (roomList.Count > 0)
                {
                    foreach (Room room1 in roomList)
                    {
                        roomTotalArea += room1.Area;

                        string programType = LightingToolMethods.GetParamVal(room1, "WW-ProgramType");
                        //Get the room level
                        Level roomLevel = room1.Level;

                        //Get the mounting height
                        double mountingHeightInput = LightingToolMethods.GetMountingHeight(roomLevel, doc);

                        //Get Room Bounding box for simpler rooms
                        BoundingBoxXYZ roomBox = room1.get_BoundingBox(null);

                        List<XYZ> roomDoorEdge = LightingToolMethods.GetRoomDoorEdge(doc, room1);
                        double roomAngle = 0;
                        int roomAngleInDeg = 0;
                        if (roomDoorEdge.Any())
                        {
                            roomAngle = LightingToolMethods.GetRoomAngle(roomDoorEdge);
                            roomAngleInDeg = LightingToolMethods.GetAngleInDeg(roomAngle);
                        }


                        if (roomAngleInDeg == 0 || roomAngleInDeg == 90 || roomAngleInDeg==180 || roomAngleInDeg == 270 || roomAngleInDeg==360)
                        {
                            using (Transaction trans = new Transaction(doc, "Pataakha: " + room1.Name))
                            {
                                trans.Start();
                                if (programType == "work")
                                {
                                    Transform tForm = null;
                                    ModelOfficeLights(light1200Symbol, roomBox, room1, doc, roomLevel, mountingHeightInput, tForm, 0);
                                }
                                trans.Commit();
                            }
                        }
                        else
                        {
                            List<XYZ> roomVertex = LightingToolMethods.GetRoomVertex(room1);
                            Transform tForm = LightingToolMethods.GetTransformObj(roomDoorEdge, -roomAngle);
                            Transform tFormBack = LightingToolMethods.GetTransformObj(roomDoorEdge, roomAngle);

                            List<XYZ> transformedRoomVertex = LightingToolMethods.TransformRoomVertex(roomVertex, tForm);
                            List<XYZ> convexHull = LightingToolMethods.GetConvexHull(transformedRoomVertex);
                            BoundingBoxXYZ transformedBB = LightingToolMethods.GetRoomWidthTransformedBB(convexHull);
                            
                            using (Transaction trans = new Transaction(doc, "Pataakha: " + room1.Name))
                            {
                                trans.Start();
                                List<FamilyInstance> lightsPlaced = new List<FamilyInstance>();
                                if (programType == "work")
                                {
                                    lightsPlaced = ModelOfficeLights(light1200Symbol, transformedBB, room1, doc, roomLevel, mountingHeightInput, tFormBack, roomAngle);
                                }

                                trans.Commit();
                            }

                        }

                    }
                }

                cleanup:
                // Return success result
                int totalLightsModeled = lightsModeled.Count();
                List<string> lightsArea = new List<string>();
                lightsArea.Add(totalLightsModeled.ToString());
                lightsArea.Add(roomTotalArea.ToString());
                string toolName = "H-1.2";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Success", doc, uiApp, detlaMilliSec, lightsArea);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                int totalLightsModeled = lightsModeled.Count();
                List<string> lightsArea = new List<string>();
                lightsArea.Add(totalLightsModeled.ToString());
                lightsArea.Add(roomTotalArea.ToString());
                string toolName = "H-1.2";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec, lightsArea);
                message = e.Message;
                return Result.Failed;
            }
        }

        public List<FamilyInstance> ModelOfficeLights
            (FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, 
            Room room1, Document doc, Element roomLevel, 
            double mountingHeight, Transform tForm, double roomAngle)
        {
            List<FamilyInstance> placedLights = new List<FamilyInstance>();
            if (!lightSymbol.IsActive)
            {
                lightSymbol.Activate();
            }
            if (roomBox != null)
            {
                Level hostLevel = roomLevel as Level;
                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;

                double L = xMaxVal - xMinVal;
                double W = yMaxVal - yMinVal;

                int mMin = (int)Math.Ceiling(W / 2700);
                //int mMax = (int)Math.Floor(W / 2100);

                int nMin = (int)Math.Ceiling(((L + 1200) / 2400) - 1);
                int nMax = (int)Math.Ceiling(((L + 1200) / 1800) - 1);

                double x;
                if (mMin == 0 || mMin < 1)
                {
                    x = 1200;
                }
                else
                {
                    x = W / (2 * mMin);
                }
                double y = (L + 1200) / (nMin + 1);

                if (nMin > 2)
                {
                    if (y < 1800)
                    {
                        y = 1800;
                    }
                    else if (y > 2400)
                    {
                        y = 2400;
                    }
                }

                for (int j = 1; j <= nMax; j++)
                {
                    for (int i = 1; i <= mMin; i++)
                    {
                        XYZ tempPoint1 = new XYZ((xMinVal + y - 600 + (j - 1) * y) / 304.8, (yMinVal + x + (i - 1) * 2 * x) / 304.8, zVal / 304.8);
                        if(tForm != null)
                        {
                            tempPoint1 = tForm.OfVector(tempPoint1);
                        }

                        XYZ tempPoint2 = new XYZ(tempPoint1.X, tempPoint1.Y, tempPoint1.Z + 3);
                        bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                        if (lightInsideRoom)
                        {
                            FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, roomLevel, hostLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            lightsModeled.Add(fi2);
                            if (tForm != null)
                            {
                                Line axis2 = Line.CreateBound(tempPoint1, tempPoint2);
                                ElementTransformUtils.RotateElement(doc, fi2.Id, axis2, roomAngle);
                            }

                            LightingToolMethods.SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            LightingToolMethods.SetMountingVal(fi2, "WW-MountingHeight", mountingHeight);
                            LightingToolMethods.ChangeOffsetToZero(fi2);
                            placedLights.Add(fi2);
                        }

                    }
                }
            }
            
            return placedLights;
        }
    }
}
