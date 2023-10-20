using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;

namespace DeskAutomation.Helper_Classes
{
    public class HelperMethods
    {

        public static List<Room> GetRoomsFromSelection(UIDocument uidoc)
        {
            //Collect the rooms from user selection, filter out unrequired elements
            List<Room> roomList = new List<Room>();
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
                        string programType = GetParamVal(testRedundant, "WW-ProgramType");

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
            return roomList;
        }
        public static String GetParamVal(Room r, string sharedParameter)
        {
            //Getting room program type parameter value
            String paraValue;
            Guid paraGuid = r.LookupParameter(sharedParameter).GUID;
            paraValue = r.get_Parameter(paraGuid).AsString().ToLower();
            return paraValue;
        }

        //Get Desk Family symbol
        public static FamilySymbol GetDeskSymbol(Document doc)
        {
            FamilySymbol deskType = null;
            string familyName = "1_person-office-desk";
            string typeName = "48x24";
            FilteredElementCollector furnitureSystem = new FilteredElementCollector(doc)
            .WhereElementIsElementType()
            .OfCategory(BuiltInCategory.OST_FurnitureSystems);

            deskType = furnitureSystem.Cast<FamilySymbol>()
            .Where(x => x.FamilyName.ToLower().Contains(familyName))
            .Where(x => x.Name.ToLower().Contains(typeName)).FirstOrDefault();

            return deskType;
        }

        //Get all doors and their rooms
        public static Dictionary<DoorData, List<Room>> GetDoors (Document doc)
        {
            Dictionary<DoorData, List<Room>> doorInfo = new Dictionary<DoorData, List<Room>>();

            FilteredElementCollector doorCollector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Doors);

            foreach (var d in doorCollector)
            {
                DoorData door = new DoorData();
                door.GetDoorInfo(d as FamilyInstance);
                doorInfo.Add(door, new List<Room> { door.EnterRoom, door.ExitRoom });
            }

            return doorInfo;
        }





        #region For rooms with <10P
        //Get the Direction vector from Room location point to Edge
        //More info: https://forums.autodesk.com/t5/revit-api-forum/find-perpendicular-point-from-one-point-to-a-line/td-p/8495308
        public static XYZ GetRoomNormal(XYZ P, XYZ Q, XYZ roomLocPt)
        {
            Line linePQ = Line.CreateBound(P, Q);
            linePQ.MakeUnbound();
            XYZ pointH = linePQ.Project(roomLocPt).XYZPoint;
            XYZ pointN = (pointH - roomLocPt).Normalize();

            return pointN;
        }
        
        //Change desk offset to zero
        public static void ChangeOffsetToZero(FamilyInstance fi)
        {
            Parameter deskOffset = fi.GetParameters("Offset").FirstOrDefault();
            deskOffset.Set(0);
        }

        public static List<FamilyInstance> PlaceDesks
            (Room room, Document doc, FamilySymbol deskType, List<XYZ> deskP, Element host, Level level, double angleToEdge)
        {
            List<FamilyInstance> deskList = new List<FamilyInstance>();
            foreach (XYZ pt in deskP)
            {
                List<XYZ> clearancePts = new List<XYZ>
                { new XYZ(pt.X, pt.Y-4.75, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-4.75, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-4.75, pt.Z + 2),
                new XYZ(pt.X, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y, pt.Z + 2)};

                Transform tForm = Transform.CreateRotationAtPoint(XYZ.BasisZ, angleToEdge, pt);
                bool clear = true;

                foreach (XYZ point in clearancePts)
                {
                    XYZ tPoint = tForm.OfPoint(point);
                    if(!room.IsPointInRoom(tPoint))
                    {
                        clear = false;
                        break;
                    }
                }


                if (clear)
                {
                    //Create Family Instance
                    FamilyInstance fi = doc.Create.NewFamilyInstance
                        (pt, deskType, host, level,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    //Rotate desk
                    XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                    Line axis = Line.CreateBound(pt, pt2);
                    ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                    ChangeOffsetToZero(fi);

                    deskList.Add(fi);
                }

            }
            return deskList;
        }

        //Get Floor below the room
        public static Element FindFloorBelow(Document doc, Room room)
        {
            /*
            Find the floor below the Room bounding box center point.
            This may fail sometimes if the room is C shaped or even longer L shaped.
            If the room bounding box center falls outside the room, we'll use Level as default reference.
            */
            Element floor = null;
            BoundingBoxXYZ roomBox = room.get_BoundingBox(null);
            XYZ center = roomBox.Min.Add(roomBox.Max).Multiply(0.5);
            if (room.IsPointInRoom(center))
            {
                int roomLevelId = room.Level.Id.IntegerValue;
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                bool isNotTemplate(View3D v3) => !(v3.IsTemplate);
                View3D view3D = collector.OfClass(typeof(View3D)).Cast<View3D>().First<View3D>(isNotTemplate);
                XYZ rayDirection = new XYZ(0, 0, -1);
                ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                try
                {
                    ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext = refIntersector.FindNearest(center, rayDirection);
                    var intersection = referenceWithContext.GetReference().ElementId;
                    floor = doc.GetElement(intersection);
                    int floorLevelId = floor.LevelId.IntegerValue;
                    if (roomLevelId != floorLevelId)
                    {
                        floor = null;
                    }
                }
                catch (Exception)
                {
                    floor = null;
                }
            }
            else
            {
                floor = null;
            }
            return floor;
        }


        //Transformation
        public static Transform GetTransformObj(List<XYZ> doorEdge, int orientation)
        {
            XYZ doorVector = (doorEdge[1] - doorEdge[0]).Normalize();
            double angleToEdge = XYZ.BasisX.AngleOnPlaneTo(doorVector, XYZ.BasisZ) * orientation;

            XYZ startPt = doorEdge[0];

            Transform tForm = Transform.CreateRotationAtPoint(XYZ.BasisZ, angleToEdge, startPt);

            return tForm;
        }


        //FOR <10P space
        public static List<XYZ> TravelFullEdge(List<XYZ> edgePts)
        {
            List<XYZ> deskPts = new List<XYZ>();

            for (int i=0; i< edgePts.Count()-1; i+=2)
            {
                deskPts.AddRange(TravelSingleEdge(edgePts[i], edgePts[i+1], "single", 5));
            }

            return deskPts;
        }

        //FOR <10P space
        public static List<XYZ> TravelSingleEdge(XYZ P, XYZ Q, string deskStyle, int deskLimit)
        {
            double deskWidth = 4;
            List<XYZ> deskP = new List<XYZ>();
            double dist = P.DistanceTo(Q);
            XYZ v = Q - P;
            XYZ vN = v.Normalize();
            //number of possible desk
            int n = (int)Math.Floor((dist - 0.25) / deskWidth);
            double remainder = (dist - 0.25) % deskWidth;
            int count = 0;


            for (int i = 1; i <= n; i++)
            {
                double d = (i - 1) * deskWidth + 2.125 + remainder;
                XYZ tPoint = P + d * vN;
                if ((count % deskLimit) == 0 && count != 0)
                {
                    count = 0;
                    continue;
                }

                deskP.Add(tPoint);
                count++;
            }

            return deskP;
        }

        #endregion


        #region Unused CODE : DO NOT DELETE
        //public static List<XYZ> TravelFullEdgeDouble(List<XYZ> edgePts, List<int> rowNumbers, string layout)
        //{
        //    List<XYZ> deskPts = new List<XYZ>();

        //    for (int i = 0; i < edgePts.Count() - 1; i += 2)
        //    {
        //        deskPts.AddRange(TravelEdgeForDoubleBank(edgePts[i], edgePts[i + 1], rowNumbers, layout));
        //    }

        //    return deskPts;
        //}


        //public static List<XYZ> TravelEdgeForDoubleBank(XYZ P, XYZ Q, List<int> rowsNumbers, string layout)
        //{
        //    double deskWidth = 4;
        //    List<XYZ> deskP = new List<XYZ>();
        //    List<XYZ> deskAll = new List<XYZ>();
        //    double dist = P.DistanceTo(Q);
        //    XYZ v = Q - P;
        //    XYZ vN = v.Normalize();
        //    //number of possible desk
        //    int n = (int)Math.Floor((dist - 0.25) / deskWidth);
        //    double remainder = (dist - 0.25) % deskWidth;

        //    for (int i = 1; i <= n; i++)
        //    {
        //        double d = (i - 1) * deskWidth + 2.125 + remainder;
        //        XYZ tPoint = P + d * vN;

        //        deskAll.Add(tPoint);
        //    }

        //    int count = 0;
        //    int firstLimit = rowsNumbers[0];
        //    if (layout == "single")
        //    {
        //        firstLimit++;
        //    }
        //    int secondLimit = rowsNumbers[1];
        //    int defaultLimit = rowsNumbers[2];

        //    int limit = firstLimit;
        //    int rowsId = 1;
        //    foreach (XYZ pt in deskAll)
        //    {
        //        if (rowsId == 1)
        //        {
        //            limit = firstLimit;
        //        }
        //        else if (rowsId == 2)
        //        {
        //            limit = secondLimit;
        //        }
        //        else
        //        {
        //            limit = defaultLimit;
        //        }

        //        if ((count % limit) == 0 && count != 0)
        //        {
        //            count = 0;
        //            rowsId++;
        //            continue;
        //        }
        //        else
        //        {
        //            deskP.Add(pt);
        //        }
        //        count++;
        //    }


        //    return deskP;
        //}


        //public static List<XYZ> TravelDoubleBanks2
        //    (string direction, RoomData roomOb)
        //{
        //    List<XYZ> doubleDesks = new List<XYZ>();

        //    double maxL = roomOb.RoomLength;
        //    double maxW = roomOb.RoomWidth;
        //    int dir = 1;
        //    XYZ lastEdgePt = roomOb.TransformedBB[3];
        //    XYZ firstEdgePt = roomOb.TransformedBB[0];


        //    if (direction == "right")
        //    {
        //        dir = -1;
        //        lastEdgePt = roomOb.TransformedBB[2];
        //        firstEdgePt = roomOb.TransformedBB[1];
        //    }


        //    double deskDepth = 4.75;

        //    double doorOffset = 4;
        //    double effWidth = maxW - 4.75;
        //    double effLength = maxL - doorOffset;//1200mm offset from door edge

        //    int m = (int)Math.Floor(effWidth / (deskDepth * 2));
        //    double mRemainder = (effWidth) % (deskDepth * 2);

        //    double lastEdgePtY = lastEdgePt.Y;
        //    double lastEdgePtZ = lastEdgePt.Z;
        //    double firstEdgePtY = firstEdgePt.Y;
        //    double firstEdgePtZ = firstEdgePt.Z;

        //    List<int> rowsNumbers = GetRowNumbers(maxL);

        //    for (int i = 1; i <= m; i++)
        //    {
        //        double d = lastEdgePt.X + dir * i * deskDepth * 2;
        //        XYZ tEndPoint = new XYZ(d, lastEdgePtY, lastEdgePtZ);

        //        XYZ tStartPoint = new XYZ(d, firstEdgePtY + doorOffset, firstEdgePtZ);

        //        doubleDesks.AddRange(TravelEdgeForDoubleBank(tStartPoint, tEndPoint, rowsNumbers, "double"));
        //    }


        //    return doubleDesks;
        //}


        //public static List<int> GetRowNumbers(double maxL)
        //{
        //    int firstRow = 4;
        //    int secondRow = 4;
        //    int lastRow = 4;

        //    switch (maxL)
        //    {
        //        case double n when (n < 24.25):
        //            firstRow = 4;
        //            break;

        //        case double n when (n >= 24.25 && n < 28.25):
        //            firstRow = 2;
        //            secondRow = 2;
        //            break;

        //        case double n when (n >= 28.25 && n < 40.25):
        //            firstRow = 3;
        //            secondRow = 4;
        //            break;

        //        case double n when (n >= 40.25 && n < 44.25):
        //            firstRow = 2;
        //            secondRow = 2;
        //            break;

        //        case double n when (n >= 44.25):
        //            firstRow = 3;
        //            secondRow = 4;
        //            break;

        //        default:
        //            firstRow = 3;
        //            secondRow = 4;
        //            break;
        //    }

        //    List<int> rowsNumbers = new List<int> { firstRow, secondRow, lastRow };
        //    return rowsNumbers;
        //}



        #endregion



        #region For rooms with >10P






        public static List<List<XYZ>> TravelDoubleBanks3
    (string direction, RoomData roomOb)
        {
            List<List<XYZ>> doubleDesks = new List<List<XYZ>>();

            double maxL = roomOb.RoomLength;
            double maxW = roomOb.RoomWidth;
            int dir = 1;
            XYZ lastEdgePt = roomOb.TransformedBB[3];
            XYZ firstEdgePt = roomOb.TransformedBB[0];


            if (direction == "right")
            {
                dir = -1;
                lastEdgePt = roomOb.TransformedBB[2];
                firstEdgePt = roomOb.TransformedBB[1];
            }


            double deskDepth = 4.75;

            double doorOffset = 4;
            double effWidth = maxW - 4.75 + 0.00656168;//2mm added tolerance
            double effLength = maxL - doorOffset;//1200mm offset from door edge

            int m = (int)Math.Floor(effWidth / (deskDepth * 2));
            double mRemainder = (effWidth) % (deskDepth * 2);

            double lastEdgePtY = lastEdgePt.Y;
            double lastEdgePtZ = lastEdgePt.Z;
            double firstEdgePtY = firstEdgePt.Y;
            double firstEdgePtZ = firstEdgePt.Z;

            for (int i = 1; i <= m; i++)
            {
                double d = lastEdgePt.X + dir * i * deskDepth * 2;
                XYZ tEndPoint = new XYZ(d, lastEdgePtY, lastEdgePtZ);

                XYZ tStartPoint = new XYZ(d, firstEdgePtY , firstEdgePtZ);

                doubleDesks.Add(TravelEdgeForDoubleBank3(tStartPoint, tEndPoint, direction));
            }


            return doubleDesks;
        }


        public static List<XYZ> TravelEdgeForDoubleBank3(XYZ P, XYZ Q, string layout)
        {
            double deskWidth = 4;
            List<XYZ> deskP = new List<XYZ>();
            double dist = P.DistanceTo(Q);
            XYZ v = Q - P;
            XYZ vN = v.Normalize();
            //number of possible desk
            int n = (int)Math.Floor((dist - 0.25) / deskWidth);
            double remainder = (dist - 0.25) % deskWidth;
            int count = 0;
            //TaskDialog.Show("Revit", "Re = " + remainder.ToString());

            for (int i = 1; i <= n; i++)
            {
                double d = (i - 1) * deskWidth + 2.125 + remainder;
                XYZ tPoint = P + d * vN;

                deskP.Add(tPoint);
                count++;
            }

            if (remainder >= 3.575 && layout == "left")
            {
                //Desk Cap
                
                XYZ lastDesk = deskP.FirstOrDefault();
                XYZ endCapPoint = new XYZ(lastDesk.X, lastDesk.Y + 2, lastDesk.Z);
                HelloRevit.endCapDesksLeft.Add(endCapPoint);
            }
            if (remainder >= 3.575 && layout == "right")
            {
                //Desk Cap

                XYZ lastDesk = deskP.FirstOrDefault();
                XYZ endCapPoint = new XYZ(lastDesk.X, lastDesk.Y + 2, lastDesk.Z);
                HelloRevit.endCapDesksRight.Add(endCapPoint);
            }

            return deskP;
        }


        public static List<List<int>> GetDeskNumbers(List<XYZ> desks, RoomData roomOb)
        {
            List<List<int>> deskN = new List<List<int>>();
            List<int> leftDeskN = new List<int>();
            List<int> rightDeskN = new List<int>();
            double leftAngle = roomOb.AngleToLeftRightEdge[0];
            double rightAngle = roomOb.AngleToLeftRightEdge[1];
            Room room = roomOb.MyRoom;
            foreach (XYZ pt in desks)
            {
                List<XYZ> clearancePts = new List<XYZ>
                { new XYZ(pt.X, pt.Y-4.75+0.00656168, pt.Z + 2),//keeping 2mm internal tolerance
                new XYZ(pt.X+2, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X, pt.Y-1, pt.Z+2),
                new XYZ(pt.X, pt.Y-3, pt.Z+2),
                new XYZ(pt.X-2, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-4.75+0.00656168, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-4.75+0.00656168, pt.Z + 2)};

                Transform tFormLeft = Transform.CreateRotationAtPoint(XYZ.BasisZ, leftAngle, pt);
                Transform tFormRight = Transform.CreateRotationAtPoint(XYZ.BasisZ, rightAngle, pt);
                bool leftPointCheck = true;
                bool rightPointCheck = true;
                foreach (XYZ point in clearancePts)
                {
                    XYZ tPointLeft = tFormLeft.OfPoint(point);

                    if (!room.IsPointInRoom(tPointLeft))
                    {
                        leftPointCheck = false;
                        break;
                    }
                }
                foreach (XYZ point in clearancePts)
                {
                    XYZ tPointRight = tFormRight.OfPoint(point);

                    if (!room.IsPointInRoom(tPointRight))
                    {
                        rightPointCheck = false;
                        break;
                    }
                }
                if (leftPointCheck)
                {
                    leftDeskN.Add(1);
                }
                else
                {
                    leftDeskN.Add(0);
                }

                if (rightPointCheck)
                {
                    rightDeskN.Add(1);
                }
                else
                {
                    rightDeskN.Add(0);
                }

            }

            deskN.Add(leftDeskN);
            deskN.Add(rightDeskN);
            return deskN;
        }

        public static List<int> GetDeskNumbersSingle(List<XYZ> desks, RoomData roomOb, string dir)
        {
            List<int> deskN = new List<int>();
            double angle = roomOb.AngleToLeftRightEdge[0];

            if (dir == "right")
            {
                angle = roomOb.AngleToLeftRightEdge[1];
            }
            Room room = roomOb.MyRoom;
            foreach (XYZ pt in desks)
            {

                List<XYZ> clearancePts = new List<XYZ>
                { new XYZ(pt.X, pt.Y-4.75+0.00656168, pt.Z + 2),//keeping 2mm internal tolerance
                new XYZ(pt.X+2, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X, pt.Y-1, pt.Z+2),
                new XYZ(pt.X, pt.Y-3, pt.Z+2),
                new XYZ(pt.X-2, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X-2, pt.Y-4.75+0.00656168, pt.Z + 2),
                new XYZ(pt.X+2, pt.Y-4.75+0.00656168, pt.Z + 2)};

                Transform tForm = Transform.CreateRotationAtPoint(XYZ.BasisZ, angle, pt);

                bool pointCheck = true;
                foreach (XYZ point in clearancePts)
                {
                    XYZ tPoint = tForm.OfPoint(point);

                    if (!room.IsPointInRoom(tPoint))
                    {
                        pointCheck = false;
                        break;
                    }
                }

                if(pointCheck)
                {
                    deskN.Add(1);
                }
                else
                {
                    deskN.Add(0);
                }
            }

            return deskN;
        }

        public static int GetTotalDeskNumber(List<int> desknum)
        {
            int totalCount = 0;

            foreach(int num in desknum)
            {
                totalCount += num;
            }

            return totalCount;
        }


        public static List<FamilyInstance> PlaceDesks2
    (Document doc, FamilySymbol deskType, List<List<XYZ>> deskP, List<List<int>> deskN,List<int> emptyRow, RoomData roomOb, double angleToEdge)
        {
            List<FamilyInstance> deskList = new List<FamilyInstance>();

            int index1 = 0;
            int index2 = deskP.Count();

            int startPoint = emptyRow[0];
            int skipPoint = emptyRow[1];
            int skipPoint2 = emptyRow[2];



            for (int i = 0; i < index2; i++)
            {
                List<XYZ> deskCol = deskP[i];
                index1 = deskCol.Count();

                for(int j = startPoint; j < index1; j++)
                {
                    if(j==skipPoint || j == skipPoint2)
                    {
                        continue;
                    }
                    if(deskN[i][j] == 1)
                    {
                        XYZ pt = deskP[i][j];
                        //Create Family Instance
                        FamilyInstance fi = doc.Create.NewFamilyInstance
                            (pt, deskType, roomOb.RoomLevelElem, roomOb.RoomLevel,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        //Rotate desk
                        XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                        Line axis = Line.CreateBound(pt, pt2);
                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                        ChangeOffsetToZero(fi);

                        deskList.Add(fi);
                    }

                }


            }

            return deskList;
        }


        public static List<FamilyInstance> PlaceDesksSingle
(Document doc, FamilySymbol deskType, List<XYZ> deskP, List<int> deskN,List<int> emptyRow, RoomData roomOb, double angleToEdge)
        {
            List<FamilyInstance> deskList = new List<FamilyInstance>();
            int count = 0;

            int startPoint = emptyRow[0];
            int skipPoint = emptyRow[1];
            int skipPoint2 = emptyRow[2];

            foreach (XYZ pt in deskP)
            {
                if(count == 0)
                {
                    //Create Family Instance
                    FamilyInstance fi = doc.Create.NewFamilyInstance
                        (pt, deskType, roomOb.RoomLevelElem, roomOb.RoomLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    //Rotate desk
                    XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                    Line axis = Line.CreateBound(pt, pt2);
                    ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                    ChangeOffsetToZero(fi);

                    deskList.Add(fi);
                    count++;

                    continue;
                }
                if (deskN[count] < 1)
                {

                }
                else if(count == skipPoint || count == skipPoint2)
                {

                }
                else
                {
                    //Create Family Instance
                    FamilyInstance fi = doc.Create.NewFamilyInstance
                        (pt, deskType, roomOb.RoomLevelElem, roomOb.RoomLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    //Rotate desk
                    XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                    Line axis = Line.CreateBound(pt, pt2);
                    ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                    ChangeOffsetToZero(fi);

                    deskList.Add(fi);
                }
                count++;
            }
            return deskList;
        }







        public static List<int> GetEmptyRows (List<List<int>> deskNo)
        {
            int maxRows = 0;

            foreach(List<int> row in deskNo)
            {
                int rowCount = 0;
                foreach(int i in row)
                {
                    if(i == 1)
                    {
                        rowCount++;
                    }
                }
                if(rowCount > maxRows)
                {
                    maxRows = rowCount;
                }
            }


            int startPoint = 0;
            int skipPoint = 0;
            int skipPoint2 = 0;
            if (maxRows < 6)
            {
                //skip first
                startPoint = 1;
            }
            else if (maxRows > 5 && maxRows < 12)
            {
                startPoint = 0;
                skipPoint = (int)Math.Ceiling((maxRows * 0.5) - 1);
                skipPoint2 = skipPoint;
            }
            else if (maxRows > 11)
            {
                startPoint = 0;
                skipPoint = (int)Math.Ceiling((maxRows * 0.3333) - 1);
                skipPoint2 = (int)Math.Ceiling((maxRows * 0.6667) - 1);
            }

            List<int> emptyDeskRow = new List<int> { startPoint, skipPoint, skipPoint2 };
            return emptyDeskRow;

        }

        

        #endregion


        public static List<XYZ> TransformList (List<XYZ> opList, Transform tFormRe)
        {
            List<XYZ> transList = new List<XYZ>();
            
            foreach(var pt in opList)
            {
                transList.Add(tFormRe.OfPoint(pt));
            }

            return transList;
        }

        public static void PrintPoints(List<XYZ> points, Transform tFormRe)
        {
            string prompt = "Printed Vertex = "
                + points.Count().ToString() + Environment.NewLine;

            foreach (var pt in points)
            {
                XYZ ptTrans = tFormRe.OfPoint(pt);
                prompt += "X = " + Math.Round(ptTrans.X).ToString()
                    + "  _&_  Y = " + Math.Round(ptTrans.Y).ToString()
                    + Environment.NewLine;
            }

            TaskDialog.Show("Revit Window: ", prompt);

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
