using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    class DimensionLights : IExternalCommand
    {
        int missedWallRef = 0;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            
            try
            {
                #region Get Rooms from Selection
                //Collect the rooms from user selection, filter out unrequired elements
                List<Room> roomList = new List<Room>();
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                

                if (0 == selectedIds.Count)
                {
                    // If no elements are selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + " You haven't selected any Rooms!");
                    goto skipTool;
                }
                else
                {
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = uidoc.Document.GetElement(id);
                        if (elem is Room)
                        {
                            Room testRedundant = elem as Room;
                            if (testRedundant.Area != 0)
                            {
                                //Add rooms to roomList
                                roomList.Add(elem as Room);
                            }
                        }
                    }
                    if (0 == roomList.Count)
                    {
                        // If no rooms are selected.
                        TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Rooms!");
                        goto skipTool;
                    }
                }
                #endregion

                foreach (Room r in roomList)
                {
                    List<FamilyInstance> roomLights = LightingToolMethods.GetLights(r);
                    if(roomLights.Count > 0)
                    {
                        List<Light> lightObjects = new List<Light>();
                        foreach (FamilyInstance fi in roomLights)
                        {
                            Light l1 = new Light();
                            l1.LightParameters(fi);
                            lightObjects.Add(l1);
                        }
                        List<List<Light>> lightXGroups = lightObjects.OrderBy(p => p.RoundedX).GroupBy(p => p.RoundedX).Select(grp => grp.ToList()).ToList();
                        List<List<Light>> lightYGroups = lightObjects.OrderBy(p => p.RoundedY).GroupBy(p => p.RoundedY).Select(grp => grp.ToList()).ToList();

                        List<Light> xLights = lightXGroups[0];
                        List<Light> yLights = lightYGroups[0];
                        ReferenceArray xRef = new ReferenceArray();
                        ReferenceArray yRef = new ReferenceArray();
                        Reference tempRef = null;

                        #region Find Y references
                        //find left Wall reference
                        tempRef = FindNextWall(doc, yLights[0], -(XYZ.BasisX));
                        if (tempRef != null)
                        {
                            yRef.Append(tempRef);
                            tempRef = null;
                        }
                        //find all lights reference
                        foreach (Light li in yLights)
                        {
                            if (li.AngleDegree == 180)
                            {
                                yRef.Append(li.LightElem.GetReferenceByName("Light Source Axis (F/B)"));
                            }
                            else
                            {
                                yRef.Append(li.LightElem.GetReferenceByName("Light Source Axis (L/R)"));
                            }
                        }
                        //find right Wall reference
                        tempRef = FindNextWall(doc, yLights[0], XYZ.BasisX);
                        if (tempRef != null)
                        {
                            yRef.Append(tempRef);
                            tempRef = null;
                        }
                        #endregion

                        #region Find X references
                        //find bottom Wall reference
                        tempRef = FindNextWall(doc, xLights[0], -(XYZ.BasisY));
                        if (tempRef != null)
                        {
                            xRef.Append(tempRef);
                            tempRef = null;
                        }
                        //find all lights reference
                        foreach (Light li in xLights)
                        {
                            if (li.AngleDegree == 180)
                            {
                                xRef.Append(li.LightElem.GetReferenceByName("Light Source Axis (L/R)"));
                            }
                            else
                            {
                                xRef.Append(li.LightElem.GetReferenceByName("Light Source Axis (F/B)"));
                            }
                        }
                        //find top Wall reference
                        tempRef = FindNextWall(doc, xLights[0], XYZ.BasisY);
                        if (tempRef != null)
                        {
                            xRef.Append(tempRef);
                            tempRef = null;
                        }
                        #endregion

                        if(missedWallRef != 0)
                        {
                            TaskDialog.Show("Revit", missedWallRef.ToString() + " Wall(s) Not Found during dimensions");
                        }
                        using (Transaction trans1 = new Transaction(doc, "Lights Dimensions"))
                        {
                            trans1.Start();
                            XYZ firstXPoint = xLights[0].XYZPoint;
                            XYZ nextXPoint = new XYZ(firstXPoint.X, firstXPoint.Y + 1, firstXPoint.Z);
                            Line vLine = Line.CreateBound(firstXPoint, nextXPoint);
                            Dimension dimenV = doc.Create.NewDimension(doc.ActiveView, vLine, xRef);
                            ElementTransformUtils.MoveElement(doc, dimenV.Id, -1 * XYZ.BasisX);

                            XYZ firstYPoint = yLights[0].XYZPoint;
                            XYZ nextYPoint = new XYZ(firstYPoint.X + 1, firstYPoint.Y, firstYPoint.Z);
                            Line hLine = Line.CreateBound(firstYPoint, nextYPoint);
                            Dimension dimenH = doc.Create.NewDimension(doc.ActiveView, hLine, yRef);
                            ElementTransformUtils.MoveElement(doc, dimenH.Id, -2 * XYZ.BasisY);
                            trans1.Commit();
                        }
                    }
                    
                }
            skipTool:

                // Return success result

                string toolName = "BulkExport";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "BulkExport";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);

                message = e.Message;
                return Result.Failed;
            }
        }

        private Reference FindNextWall(Document doc, Light fi, XYZ rayDirection)
        {
            /*
            Find the nearest wall to the end fixtures for placing directions.
            */
            Reference reference = null;

            XYZ lightPoint = fi.XYZPoint;
            XYZ offsetLightPoint = new XYZ(lightPoint.X, lightPoint.Y, lightPoint.Z + 1);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            bool isNotTemplate(View3D v3) => !(v3.IsTemplate);
            View3D view3D = collector.OfClass(typeof(View3D)).Cast<View3D>().First<View3D>(isNotTemplate);
            ElementClassFilter filter = new ElementClassFilter(typeof(Wall));
            try
            {
                ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face, view3D);
                ReferenceWithContext referenceWithContext = refIntersector.FindNearest(offsetLightPoint, rayDirection);
                reference = referenceWithContext.GetReference();
                var intersection = reference.ElementId;
                var wall = doc.GetElement(intersection);
                var wallCurve = wall.Location as LocationCurve;
                double dist = wallCurve.Curve.Distance(offsetLightPoint);
                //if (dist *304.8 > 2800)
                //{
                //    reference = null;
                //}
            }
            catch (Exception)
            {
                missedWallRef++;
            }

            return reference;
        }
    }


    public class Light
    {
        // Class for creating light as objects and create properties related to light family instance

        public void LightParameters(FamilyInstance light)
        {
            LightElem = light;
            Location = light.Location as LocationPoint;
            XYZPoint = Location.Point;
            XYZPoint = light.GetSpatialElementCalculationPoint();
            LightOrientation = light.FacingOrientation;
            AngleDegree = Math.Round(LightOrientation.AngleOnPlaneTo(XYZ.BasisX, XYZ.BasisZ)* 180 / Math.PI);
            RoundedX = Math.Round(Location.Point.X);
            RoundedY = Math.Round(Location.Point.Y);
        }

        public FamilyInstance LightElem { get; set; } //Light as a family instance
        public LocationPoint Location { get; set; } //Location of the Light
        public XYZ XYZPoint { get; set; } //XYZ point of Light
        public XYZ LightOrientation { get; set; } //Facing Orientation of Light
        public double RoundedX { get; set; } //X coordinate of Light rounded to 1 foot
        public double RoundedY { get; set; } //Y coordinate of Light rounded to 1 foot
        public double AngleDegree { get; set; } //Facing angle of Light in Degrees rounded to 1 degree

    }
}

