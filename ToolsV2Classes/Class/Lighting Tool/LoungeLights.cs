using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using BoundarySegment = Autodesk.Revit.DB.BoundarySegment;



namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class LoungeLights : IExternalCommand
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
                //Get Lighting fixture family Type
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                FamilySymbol lightCanSymbol = collector.OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(x => x.FamilyName.Contains("IN-Architecture Point-07"))
                    .FirstOrDefault(x => x.Name.Contains("WWI-LT-AP-07-01"));
                //Check whether lighting fixture family exists in Project, if not try to load from Google drive
                if (lightCanSymbol == null)
                {
                    TaskDialog.Show("Revit", "Standard Lighting fixture family is not loaded into the project."
                        + Environment.NewLine
                        + "Tool will try to load the latest lighting fixture Family for cylindrical lights: IN-Architecture Point-07");
                    //Loading family from G drive
                    using (Transaction tx = new Transaction(doc, "Load Light Fixture Family"))
                    {
                        tx.Start();
                        if (lightCanSymbol == null)
                        {
                            string path = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Light Fixture\\IN-Architecture Point-07.rfa";
                            //Check familyLoadOption subClass in LoungeLights class
                            FamilyLoadOption newOption = new FamilyLoadOption();
                            doc.LoadFamily(path, newOption, out Family lightFamily);
                            if(lightFamily == null)
                            {
                                //error in loading latest family
                                TaskDialog.Show("Revit", "Light Family cannot be Loaded, please load the family 'IN-Architecture Point-07' manually :(");
                            }
                            else
                            {
                                //Family loaded successfully
                                TaskDialog.Show("Revit", "Light Family Loaded:" + lightFamily.Name + Environment.NewLine + "Please run the tool again :)");
                            }
                        }
                        tx.Commit();
                    }
                    goto cleanup;
                }
                #endregion

                //Collect the rooms from user selection, filter out unrequired elements
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
                            if (testRedundant.Area != 0)
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
                    if (!(null == lightCanSymbol))
                    {
                        foreach (Room room1 in roomList)
                        {
                            roomTotalArea += room1.Area;
                            //Get room program type parameter
                            string programType = LightingToolMethods.GetParamVal(room1, "WW-ProgramType");

                            //Get the room level
                            Level roomLevel = room1.Level;

                            //Get the mounting height
                            double mountingHeightInput = LightingToolMethods.GetMountingHeight(roomLevel, doc);

                            //Get Room Bounding box
                            BoundingBoxXYZ roomBox = room1.get_BoundingBox(null);

                            if(programType == "we")
                            {
                                using (Transaction trans = new Transaction(doc, "Pataakha: " + room1.Name))
                                {
                                    trans.Start();
                                    ModelLoungeLights(lightCanSymbol, roomBox, room1, doc, roomLevel, mountingHeightInput);
                                    trans.Commit();
                                }
                            }

                            if (programType == "circulate")
                            {
                                //Start transaction for creating lighting fixtures in each selected room
                                using (Transaction trans = new Transaction(doc, "Pataakha: " + room1.Name))
                                {
                                    trans.Start();

                                    //Check ModelCorridorLights method
                                    ModelCorridorLights(lightCanSymbol, roomBox, room1, doc, roomLevel, mountingHeightInput);
                                    trans.Commit();
                                }
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
                string toolName = "you CAN";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Success", doc, uiApp, detlaMilliSec, lightsArea);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                int totalLightsModeled = lightsModeled.Count();
                List<string> lightsArea = new List<string>();
                lightsArea.Add(totalLightsModeled.ToString());
                lightsArea.Add(roomTotalArea.ToString());
                string toolName = "you CAN";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec, lightsArea);
                message = e.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Model Lights for Lounges or any spaces with "We" as program type
        /// </summary>
        /// <param name="lightSymbol"></param>
        /// <param name="roomBox"></param>
        /// <param name="room1"></param>
        /// <param name="doc"></param>
        /// <param name="roomLevel"></param>
        /// <param name="mountingHeight"></param>
        public void ModelLoungeLights(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc,
            Element roomLevel, double mountingHeight)
        {
            if (!lightSymbol.IsActive)
            {
                //Activate symbol
                lightSymbol.Activate();
            }
            if (roomBox != null)
            {
                Level hostLevel = room1.Level as Level;
                //Get min max value of room bounding box

                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;

                //Get length and width of room bounding box
                double L = (xMaxVal - xMinVal);
                double W = (yMaxVal - yMinVal);
                if (L > 1500 && W > 1500)
                {
                    //Calculate minimum and maximum number of lights possible based on standard spacing
                    //Currently we are optimising in such a way that the spacing is maximum and number of lights are minimum
                    //This nMin is for horizontal lights, mMin is for vertical lights
                    int nMin = (int)Math.Ceiling((L - 2400 + 2200) / 2200);
                    //int nMax = (int)Math.Floor((L - 2400 + 1500) / 1500);

                    int mMin = (int)Math.Ceiling((W - 2400 + 2200) / 2200);
                    //int mMax = (int)Math.Floor((W - 2400 + 1500) / 1500);
                    double x;
                    double y;
                    //For smaller rooms, we need to do some approximations as below
                    if (nMin < 2)
                    {
                        x = 1800;
                    }
                    else
                    {
                        x = (L - 2400) / (nMin - 1);
                    }
                    //For smaller rooms, we need to do some approximations as below
                    if (mMin < 2)
                    {
                        y = 1800;
                    }
                    else
                    {
                        y = (W - 2400) / (mMin - 1);
                    }
                    //Creating lights based on above variables
                    for (int j = 1; j <= mMin; j++)
                    {
                        for (int i = 1; i <= nMin; i++)
                        {
                            XYZ tempPoint1 = new XYZ((xMinVal + 1200 + (i - 1) * x) / 304.8, (yMinVal + 1200 + (j - 1) * y) / 304.8, zVal / 304.8);
                            XYZ tempPoint2 = new XYZ((xMinVal + 1200 + (i - 1) * x) / 304.8, (yMinVal + 1200 + (j - 1) * y) / 304.8, (zVal + 1000) / 304.8);
                            //Checking whether center of light is falling inside room
                            bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                            if (lightInsideRoom)

                            {
                                //Modelling the light fixture
                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, roomLevel, hostLevel, StructuralType.NonStructural);
                                lightsModeled.Add(fi2);
                                //Setting bottom of fixture and mounting height parameter of Lights 
                                LightingToolMethods.SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                                LightingToolMethods.SetMountingVal(fi2, "WW-MountingHeight", mountingHeight);
                                LightingToolMethods.ChangeOffsetToZero(fi2);
                            }
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Model Lights for Lounges or any spaces with "Cuirculate" as program type
        /// </summary>
        /// <param name="lightSymbol"></param>
        /// <param name="roomBox"></param>
        /// <param name="room1"></param>
        /// <param name="doc"></param>
        /// <param name="roomLevel"></param>
        /// <param name="mountingHeight"></param>
        public void ModelCorridorLights(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc, Element roomLevel, double mountingHeight)
        {
            if (!lightSymbol.IsActive)
            {
                //Activate symbol
                lightSymbol.Activate();
            }
            if (roomBox != null)
            {
                Level hostLevel = roomLevel as Level;
                //Get min max value of room bounding box
                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;
                //Get length and width of room bounding box
                double L = (xMaxVal - xMinVal);
                double W = (yMaxVal - yMinVal);

                //For horizontal corridors
                if (L>W)
                {
                    //Calculate minimum and maximum number of lights possible based on standard spacing
                    //Currently we are optimising in such a way that the spacing is maximum and number of lights are minimum
                    //This nMin is for horizontal lights
                    int nMin = (int)Math.Ceiling((L - W + 2100) / 2100);
                    //int nMax = (int)Math.Floor((L - W + 1500) / 1500);

                    double x;
                    double y;
                    //Approximation for smaller corridor as below
                    if (nMin > 1)
                    {
                        x = (L - W) / (nMin - 1);
                    }
                    else
                    {
                        x = L / 2;
                    }
                    
                    y = W / 2;
                    //Creating lights based on above variables
                    for (int i = 1; i <= nMin+1; i++)
                    {
                        XYZ tempPoint1 = new XYZ((xMinVal + (W*0.5) + (i-1) * x) / 304.8, (yMinVal + y) / 304.8, zVal / 304.8);
                        XYZ tempPoint2 = new XYZ((xMinVal + (W * 0.5) + (i - 1) * x) / 304.8, (yMinVal + y) / 304.8, (zVal + 1000) / 304.8);
                        //Checking whether center of light is falling inside room
                        bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                        if (lightInsideRoom)
                        {
                            //Modelling the light fixture
                            FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, roomLevel, hostLevel, StructuralType.NonStructural);
                            lightsModeled.Add(fi2);
                            //Setting bottom of fixture and mounting height parameter of Lights
                            LightingToolMethods.SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            LightingToolMethods.SetMountingVal(fi2, "WW-MountingHeight", mountingHeight);
                            LightingToolMethods.ChangeOffsetToZero(fi2);
                        }
                    }
                }
                //For Vertical corridors
                else if (W >= L)
                {
                    //Calculate minimum and maximum number of lights possible based on standard spacing
                    //Currently we are optimising in such a way that the spacing is maximum and number of lights are minimum
                    //This nMin is for vertical lights
                    int nMin = (int)Math.Ceiling((W - L + 2100) / 2100);
                    //int nMax = (int)Math.Floor((W - L + 1500) / 1500);

                    //Approximation for smaller corridor as below
                    double x;
                    double y;
                    if (nMin > 1)
                    {
                        y = (W - L) / (nMin -1);
                    }
                    else
                    {
                        y = W / 2;
                    }
                    x = L / 2;
                    //Creating lights based on above variables
                    for (int i = 1; i <= nMin; i++)
                    {
                        XYZ tempPoint1 = new XYZ((xMinVal + x) / 304.8, (yMinVal + L*0.5 + (i-1)*y) / 304.8, zVal / 304.8);
                        XYZ tempPoint2 = new XYZ((xMinVal + x) / 304.8, (yMinVal + L * 0.5 + (i - 1) * y) / 304.8, (zVal + 1000) / 304.8);
                        //Checking whether center of light is falling inside room
                        bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                        if (lightInsideRoom)
                        {
                            //Modelling the light fixture
                            FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, roomLevel, hostLevel, StructuralType.NonStructural);
                            lightsModeled.Add(fi2);
                            //Setting bottom of fixture and mounting height parameter of Lights
                            LightingToolMethods.SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            LightingToolMethods.SetMountingVal(fi2, "WW-MountingHeight", mountingHeight);
                            LightingToolMethods.ChangeOffsetToZero(fi2);
                        }
                    }
                }
            }
        }

        //Commented out code due to incomplete logic
        #region Irregular Corridors
        //public List<XYZ> ReturnVertexCenters (Room corridorRoom)
        //{
        //    SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
        //    {
        //        SpatialElementBoundaryLocation =
        //      SpatialElementBoundaryLocation.Finish
        //    };

        //    IList<IList<BoundarySegment>> loops = corridorRoom.GetBoundarySegments(opt);

        //    List<XYZ> roomVertices = new List<XYZ>(); //List of all room vertices
        //    foreach (IList<BoundarySegment> loop in loops)
        //    {
        //        //TaskDialog.Show("Revit", "Total Segments = " + loop.Count().ToString());

        //        XYZ p0 = null; //previous segment start point
        //        XYZ p = null; // segment start point
        //        XYZ q = null; // segment end point

        //        foreach (BoundarySegment seg in loop)
        //        {
        //            q = seg.GetCurve().GetEndPoint(1);

        //            if (p == null)
        //            {
        //                roomVertices.Add(seg.GetCurve().GetEndPoint(0));
        //                p = seg.GetCurve().GetEndPoint(0);
        //                p0 = p;
        //                continue;
        //            }
        //            p = seg.GetCurve().GetEndPoint(0);
        //            if (p != null && p0 != null)
        //            {
        //                if (AreCollinear(p0, p, q))//skipping the segments that are collinear
        //                {
        //                    p0 = p;
        //                    continue;
        //                }
        //                else
        //                {
        //                    roomVertices.Add(p);
        //                }
        //            }
        //            p0 = p;
        //        }
        //    }

        //    double tolerance = 2550; //Distance between two Points (in mm) should be less than this number
        //    List<List<XYZ>> nearbyPairs = new List<List<XYZ>>(); //List of Pairs of nearby points
        //    for (int i = 0; i < roomVertices.Count() - 1; i++)
        //    {
        //        for (int j = i + 1; j < roomVertices.Count(); j++)
        //        {
        //            double dist = roomVertices[i].DistanceTo(roomVertices[j]) * 304.8;
        //            if (dist < tolerance) //checking whether two points are nearby based on tolerance
        //            {
        //                nearbyPairs.Add(new List<XYZ> { roomVertices[i], roomVertices[j] });
        //            }
        //        }
        //    }
        //    //TaskDialog.Show("Revit", "Total points = " + roomVertices.Count().ToString()
        //    //    + Environment.NewLine + "Total Pairs = " + nearbyPairs.Count());

        //    List<XYZ> centerPoints = new List<XYZ>();
        //    foreach (List<XYZ> pair in nearbyPairs)
        //    {
        //        centerPoints.Add((pair[0] + pair[1]) / 2);
        //    }

        //    return centerPoints;


        //}

        //static bool AreCollinear (XYZ p1, XYZ p2, XYZ p3)
        //{
        //    bool collinear = false;
        //    double area = 0.5*Math.Abs(p1.X * (p2.Y - p3.Y)
        //        + p2.X * (p3.Y - p1.Y)
        //        + p3.X * (p1.Y - p2.Y));
        //    //sometimes area is not exactly zero but is very small number
        //    if (area < 0.1)
        //    {
        //        collinear = true;
        //    }
        //    return collinear;
        //}

        //public void CorridorLightsNew(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc, Element roomLevel, double mountingHeight)
        //{
        //    List<XYZ> centerPoints = ReturnVertexCenters(room1);
        //    List<List<XYZ>> sortedX = centerPoints.OrderBy(p => p.X).GroupBy(p => p.X).Select(grp => grp.ToList()).ToList();
        //    List<List<XYZ>> sortedY = centerPoints.OrderBy(p => p.Y).GroupBy(p => p.Y).Select(grp => grp.ToList()).ToList();

        //    foreach (var xGrp in sortedX)
        //    {
        //        if(xGrp.Count() >2)
        //        {

        //        }
        //    }

        //}
        #endregion

    }

    public class FamilyLoadOption : IFamilyLoadOptions
    {
        //For loading and overwriting families as per Revit api documentation
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }

    public class LightingToolMethods
    {
        //For repeatative methods in lighting tool classes
        public static String GetParamVal(Room r, string sharedParameter)
        {
            //Getting room program type parameter value
            String paraValue;
            Guid paraGuid = r.LookupParameter(sharedParameter).GUID;
            paraValue = r.get_Parameter(paraGuid).AsString().ToLower();
            return paraValue;
        }

        public static void SetMountingLinear (double mountingHeight, FamilyInstance light)
        {
            Parameter mountHeightLinear = light.GetParameters("MountingHeight").FirstOrDefault();
            mountHeightLinear.Set(mountingHeight / 304.8);

        }
        public static void SetBOLFVal(FamilyInstance fi, string sharedParameter, double BOFH)
        {
            //Setting bottom of fixture height parameter
            Parameter para = fi.LookupParameter(sharedParameter);
            if (!para.IsReadOnly)
            {
                para.Set(BOFH / 304.8);
            }
        }

        public static void SetMountingVal(FamilyInstance fi, string sharedParameter, double mountingHeight)
        {
            //Setting Mounting height parameter
            Parameter para = fi.LookupParameter(sharedParameter);
            if (!para.IsReadOnly)
            {
                para.Set(mountingHeight / 304.8);
            }
        }

        public static double GetMountingHeight(Level roomLevel, Document doc)
        {
            Level levelAbove = roomLevel;
            Level tempLevel = roomLevel;
            double zValue = roomLevel.Elevation;
            double mountingHeight = 3000;
            FilteredElementCollector allLevels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType();
            List<Level> levelList = allLevels.OfType<Level>().OrderBy(lev => lev.Elevation).ToList();
            foreach (Level e in levelList)
            {
                tempLevel = e;
                if (tempLevel.Elevation > roomLevel.Elevation && tempLevel.Elevation - roomLevel.Elevation > 2600/304.8)
                {
                    levelAbove = tempLevel;
                    goto loopbreak;
                }
            }
            //Set the mounting height value based on level above or default to 3850mm
            loopbreak:
            if (levelAbove == roomLevel)
            {
                mountingHeight = 3000;
            }
            else
            {
                mountingHeight = (levelAbove.Elevation * 304.8 - roomLevel.Elevation * 304.8) - 200;
            }
            if (mountingHeight > 6000)
            {
                //for Failsafe
                mountingHeight = 3000;
            }

            return mountingHeight;
        }

        public static void ChangeOffsetToZero(FamilyInstance fi2)
        {
            Parameter lightOffset = fi2.GetParameters("Elevation from Level").FirstOrDefault();


            if (lightOffset == null)
            {
                lightOffset = fi2.GetParameters("Offset").FirstOrDefault();
            }
            lightOffset.Set(0);
        }

        public static List<FamilyInstance> GetLights(Room room)
        {
            /*             
            Returns list of lights family instances from selected room.
            */
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter filter
              = new BoundingBoxIntersectsFilter(outline);
            Document doc = room.Document;
            FilteredElementCollector familyInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_LightingFixtures)
                .WherePasses(filter);
            int roomid = room.Id.IntegerValue;
            List<FamilyInstance> a = new List<FamilyInstance>();
            foreach (FamilyInstance fi in familyInstances)
            {
                if (null != fi.Room
                    && fi.Room.Id.IntegerValue.Equals(roomid)
                    && fi.Symbol.FamilyName.ToLower().Contains("light"))
                {
                    a.Add(fi);
                }
            }
            return a;
        }


        //Get element id of the wall with Door in it
        public static List<XYZ> GetRoomDoorEdge(Document doc, Room room)
        {
            List<XYZ> doorEdgeEndpoints = new List<XYZ>();

            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);

            foreach (IList<BoundarySegment> loop in loops)
            {
                foreach (BoundarySegment seg in loop)
                {
                    if (null != seg.ElementId)
                    {
                        ElementId edgeId = seg.ElementId;
                        Element edgeElem = doc.GetElement(edgeId);
                        if (edgeElem != null)
                        {
                            if (edgeElem.Category.Name.Equals("Walls"))
                            {
                                Wall edgeWall = edgeElem as Wall;
                                IList<ElementId> inserts =
                                    edgeWall.FindInserts(true, true, true, true);

                                if (inserts.Count() > 0)
                                {
                                    var doorEdgeVertices = seg.GetCurve().Tessellate();
                                    doorEdgeEndpoints = doorEdgeVertices as List<XYZ>;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return doorEdgeEndpoints;
        }

        public static double GetRoomAngle(List<XYZ> doorEdge)
        {
            XYZ doorVector = (doorEdge[1] - doorEdge[0]).Normalize();
            double angleToEdge = XYZ.BasisX.AngleOnPlaneTo(doorVector, XYZ.BasisZ);
            return angleToEdge;
        }

        public static int GetAngleInDeg (double angleToEdge)
        {
            int angleInDeg = (int)Math.Round(angleToEdge * 180 / Math.PI);
            return angleInDeg;
        }

        //Get all the vertices of room, I mean ALL even rooms with holes.
        //TODO: Group vertices with different edge loops
        public static List<XYZ> GetRoomVertex(Room room)
        {
            List<XYZ> roomVertices = new List<XYZ>(); //List of all room vertices
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);


            foreach (IList<BoundarySegment> loop in loops)
            {
                XYZ p0 = null; //previous segment start point
                XYZ p = null; // segment start point
                XYZ q = null; // segment end point

                foreach (BoundarySegment seg in loop)
                {
                    q = seg.GetCurve().GetEndPoint(1);

                    if (p == null)
                    {
                        roomVertices.Add(seg.GetCurve().GetEndPoint(0));
                        p = seg.GetCurve().GetEndPoint(0);
                        p0 = p;
                        continue;
                    }
                    p = seg.GetCurve().GetEndPoint(0);
                    if (p != null && p0 != null)
                    {
                        if (AreCollinear(p0, p, q))//skipping the segments that are collinear
                        {
                            p0 = p;
                            continue;
                        }
                        else
                        {
                            roomVertices.Add(p);
                        }
                    }
                    p0 = p;
                }
            }
            return roomVertices;
        }

        //Helper Method_Check whether three points are collinear
        public static bool AreCollinear(XYZ p1, XYZ p2, XYZ p3)
        {
            bool collinear = false;
            double area = 0.5 * Math.Abs(p1.X * (p2.Y - p3.Y)
                + p2.X * (p3.Y - p1.Y)
                + p3.X * (p1.Y - p2.Y));
            //sometimes area is not exactly zero but is very small number
            if (area < 0.1)
            {
                collinear = true;
            }
            return collinear;
        }




        public static Transform GetTransformObj(List<XYZ> doorEdge, double angleToEdge)
        {
            XYZ startPt = doorEdge[0];
            XYZ offsetPt = new XYZ(startPt.X, startPt.Y, startPt.Z + 1);
            XYZ axis = (offsetPt - startPt).Normalize();

            Transform tForm = Transform.CreateRotationAtPoint(axis, angleToEdge, startPt);


            return tForm;
        }


        //Transform Room vertices to get horizontal door
        public static List<XYZ> TransformRoomVertex(List<XYZ> vertexList, Transform tForm)
        {
            List<XYZ> transRoomVertex = new List<XYZ>();

            foreach (XYZ pt in vertexList)
            {
                XYZ rotatedPt = tForm.OfVector(pt);
                transRoomVertex.Add(rotatedPt);
            }
            return transRoomVertex;
        }

        //Convex H U L L
        public static List<XYZ> GetConvexHull(List<XYZ> points)
        {
            if (points == null) throw new ArgumentNullException(nameof(points));
            XYZ startPoint = points.MinBy(p => p.X);
            var convexHullPoints = new List<XYZ>();
            XYZ walkingPoint = startPoint;
            XYZ refVector = XYZ.BasisY.Negate();
            do
            {
                convexHullPoints.Add(walkingPoint);
                //setup WP and RV
                XYZ wp = walkingPoint;
                XYZ rv = refVector;
                //Find the nextWP with minimum angle between rv and (p-wp)
                walkingPoint = points.MinBy(p =>
                {
                    double angle = (p - wp).AngleOnPlaneTo(rv, XYZ.BasisZ);
                    if (angle < 1e-10) angle = 2 * Math.PI;
                    return angle;
                });
                refVector = wp - walkingPoint;
            } while (walkingPoint != startPoint);
            convexHullPoints.Reverse();
            return convexHullPoints;
        }

        //Get Room width from Transformed bounding Box based on Door edge
        public static BoundingBoxXYZ GetRoomWidthTransformedBB(List<XYZ> vertexList)
        {
            double xMin = (vertexList.OrderBy(p => p.X).FirstOrDefault()).X;
            double xMax = (vertexList.OrderBy(p => p.X).LastOrDefault()).X;
            double yMin = (vertexList.OrderBy(p => p.Y).FirstOrDefault()).Y;
            double yMax = (vertexList.OrderBy(p => p.Y).LastOrDefault()).Y;
            double z = vertexList[0].Z;
            XYZ P = new XYZ(xMin, yMin, z);
            XYZ R = new XYZ(xMax, yMax, z);

            BoundingBoxXYZ transBB = new BoundingBoxXYZ
            {
                Min = P,
                Max = R
            };

            return transBB;

        }




    }

    public static class IEnumerableExtensions
    {
        public static tsource MinBy<tsource, tkey>(
          this IEnumerable<tsource> source,
          Func<tsource, tkey> selector)
        {
            return source.MinBy(selector, Comparer<tkey>.Default);
        }
        public static tsource MinBy<tsource, tkey>(
          this IEnumerable<tsource> source,
          Func<tsource, tkey> selector,
          IComparer<tkey> comparer)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (selector == null) throw new ArgumentNullException(nameof(selector));
            if (comparer == null) throw new ArgumentNullException(nameof(comparer));
            using (IEnumerator<tsource> sourceIterator = source.GetEnumerator())
            {
                if (!sourceIterator.MoveNext())
                    throw new InvalidOperationException("Sequence was empty");
                tsource min = sourceIterator.Current;
                tkey minKey = selector(min);
                while (sourceIterator.MoveNext())
                {
                    tsource candidate = sourceIterator.Current;
                    tkey candidateProjected = selector(candidate);
                    if (comparer.Compare(candidateProjected, minKey) < 0)
                    {
                        min = candidate;
                        minKey = candidateProjected;
                    }
                }
                return min;
            }
        }
    }

}

