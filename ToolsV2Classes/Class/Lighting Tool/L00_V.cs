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
    public class L00_V : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;

            try
            {
                #region Get Lighting Fixture Family
                //Get Lighting fixture Type for Office Space
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                FamilySymbol opsLightSymbol = collector.OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_LightingFixtures)
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(x => x.FamilyName.Contains("IN-Surface Mounted Light-07"))
                    .FirstOrDefault(x => x.Name.Contains("WWI-LT-LM-07")
                    || x.Name.Contains("WWI-LT-LM07"));

                if (opsLightSymbol == null)
                {
                    TaskDialog.Show("Revit", "Standard Lighting fixture family is not loaded into the project."
                        + Environment.NewLine
                        + "Tool will try to load the latest lighting fixture Family for linear lights: IN-Surface Mounted Light-07");

                    using (Transaction tx = new Transaction(doc, "Load Light Fixture Family"))
                    {
                        tx.Start();
                        if (opsLightSymbol == null)
                        {
                            string path = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Light Fixture\\IN-Surface Mounted Light-07.rfa";
                            FamilyLoadOption newOption = new FamilyLoadOption();
                            doc.LoadFamily(path, newOption, out Family lightFamily);
                            if (lightFamily == null)
                            {
                                TaskDialog.Show("Revit", "Light Family cannot be Loaded, please load the family 'IN-Surface Mounted Light-07' manually :(");
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
                            if (testRedundant.Area != 0 && programType.ToLower().Contains("operate"))
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
                        string programType = LightingToolMethods.GetParamVal(room1, "WW-ProgramType");
                        //Get the room level
                        Level roomLevel = room1.Level;

                        //Get the mounting height
                        double mountingHeightInput = LightingToolMethods.GetMountingHeight(roomLevel, doc);

                        //Get Room Bounding box for simpler rooms
                        BoundingBoxXYZ roomBox = room1.get_BoundingBox(null);
                        using (Transaction trans = new Transaction(doc, "Create Lighting Fixtures"))
                        {
                            trans.Start();
                            if (programType == "operate")
                            {
                                ModelOpsLights(opsLightSymbol, roomBox, room1, doc, roomLevel, mountingHeightInput);
                            }
                            trans.Commit();
                        }
                    }
                }

            cleanup:
                // Return success result

                string toolName = "V-L00";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "V-L00";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }
        }
        public void ModelOpsLights(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc, Element roomLevel, double mountingHeight)
        {
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

                int nMin = (int)(L / 1200);
                //int mMax = (int)Math.Floor(W / 2100);

                int mMin = (int)((W + 600) / 3000);
                //int nMax = (int)Math.Ceiling(((L + 2400) / 3000) - 1);

                double x;
                if (nMin == 0 || nMin < 1)
                {
                    x = L / 2;
                    nMin = 1;
                }
                else
                {
                    x = (L - (nMin - 1) * 1200) * 0.5;
                }

                double y;
                if (mMin ==0 || mMin <1)
                {
                    y = W * 0.5;
                    mMin = 1;
                }
                else
                {
                    y = (W - 3000 * (mMin - 1)) * 0.5;
                }

                for (int j = 1; j <= mMin; j++)
                {
                    for (int i = 1; i <= nMin; i++)
                    {
                        XYZ tempPoint1 = new XYZ((xMinVal + x + (i - 1) * 1200) / 304.8, (yMinVal + y + (j - 1) * 3000) / 304.8, zVal / 304.8);
                        XYZ tempPoint2 = new XYZ((xMinVal + x + (i - 1) * 1200) / 304.8, (yMinVal + y + (j - 1) * 3000) / 304.8, (zVal + 1000) / 304.8);
                        bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                        if (lightInsideRoom)
                        {
                            FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, roomLevel, hostLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            Line axis2 = Line.CreateBound(tempPoint1, tempPoint2);
                            ElementTransformUtils.RotateElement(doc, fi2.Id, axis2, Math.PI / 2);
                            //LightingToolMethods.SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            LightingToolMethods.SetMountingVal(fi2, "WW-MountingHeight", mountingHeight);
                            LightingToolMethods.ChangeOffsetToZero(fi2);
                        }
                    }
                }
            }
        }
    }
}
