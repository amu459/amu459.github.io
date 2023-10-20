using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using DeskAutomation.Helper_Classes;

namespace DeskAutomation.Helper_Classes
{
    public class RoomProgramData
    {
        public Room MyRoom { get; set; } //Room
        public Element RoomLevelElem { get; set; } //Room Level as an Element
        public Level RoomLevel { get; set; } //Room Level as a Level
        public IList<IList<BoundarySegment>> Loops { get; set; } //Boundary segment loops
        public XYZ RoomLocationPoint { get; set; } //Room Location Point
        public Transform RoomTransformObj { get; set; } //Transform of the room vertices wrt door edge angle
        public List<XYZ> RoomVertex { get; set; } //Room vertices

        public string RoomOrientation { get; set; } //Horizontal or vertical Desk orientation

        public List<XYZ> LeftEdge { get; set; } //Room Vertices of extreme left edge of door
        public List<XYZ> RightEdge { get; set; } //Room Vertices of extreme right edge of door

        public double RoomAngle { get; set; } //angle of the door which is considered as angle of the room itself
        public double RoomLeftAngle { get; set; } //angle of the door which is considered as angle of the room itself
        public double RoomRightAngle { get; set; } //angle of the door which is considered as angle of the room itself

        //public List<XYZ> RoomConvexHull { get; set; } //Room convex hull derived from Room vertices

        public double RoomWidth { get; set; } //Room width
        public double RoomLength { get; set; } //Room Length
        public string RoomType { get; set; } //To decide which algorithm to use
        public List<XYZ> TransformedBB { get; set; }
        public BoundingBoxXYZ RoomBB { get; set; }




        //For Angled Rooms
        public List<XYZ> TransRoomVertex { get; set; } //Transformed Room Vertices






        public void GetRoomInfo(Document doc, Room room, string orientation)
        {
            MyRoom = room;
            RoomLevel = room.Level;
            RoomLevelElem = room.Level as Element;
            Loops = GetBoundaryLoops(room);
            RoomLocationPoint = ((LocationPoint)room.Location).Point;
            RoomOrientation = orientation;
            if(RoomOrientation == "vertical")
            {
                RoomAngle = Math.PI * 0.5;
                RoomLeftAngle = Math.PI * 0.5;
                RoomRightAngle = Math.PI * 1.5;
            }
            else if(RoomOrientation == "horizontal")
            {
                RoomAngle = 0;
                RoomLeftAngle = 0;
                RoomRightAngle = Math.PI;
            }


            //Find floor below
            Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
            if (floorBelow != null)
            {
                RoomLevelElem = floorBelow;
            }

            //vertices of rooms irrespective of internal holes
            RoomVertex = GetRoomVertex(Loops);

            RoomBB = room.get_BoundingBox(null);

            GetRoomDimensions(room, this);


            if (RoomWidth < 14.25 && RoomWidth > 5.5)
            {
                RoomType = "LeftRightSingle";
            }
            if (RoomWidth > 14.25)
            {
                RoomType = "Double";
            }

        }


        public void GetAngledRoomInfo(Document doc, Room room, string orientation)
        {
            MyRoom = room;
            RoomLevel = room.Level;
            RoomLevelElem = room.Level as Element;
            Loops = GetBoundaryLoops(room);
            RoomLocationPoint = ((LocationPoint)room.Location).Point;
            RoomOrientation = "vertical";
            double angleInDeg = 0;
            bool success = double.TryParse(orientation, out angleInDeg);
            if(success)
            {
                RoomAngle = angleInDeg * Math.PI / 180;
                RoomLeftAngle =  Math.PI * 0.5 - RoomAngle;
                RoomRightAngle = -Math.PI * 0.5 - RoomAngle;
                //RoomLeftAngle = 0;
                //RoomRightAngle = Math.PI;
            }



            ////Find floor below
            //Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
            //if (floorBelow != null)
            //{
            //    RoomLevelElem = floorBelow;
            //}

            //vertices of rooms irrespective of internal holes
            RoomVertex = GetRoomVertex(Loops);


            //Transform Room Vertex
            //TaskDialog.Show("Origin", "Room Origin = " + RoomVertex[0].X.ToString() + Environment.NewLine
            //    + RoomVertex[0].Y.ToString());
            RoomTransformObj = Transform.CreateRotationAtPoint(XYZ.BasisZ, RoomAngle, RoomVertex[0]);
            TransRoomVertex = TransformRoomVertex(RoomVertex, RoomTransformObj);

            //Transformed room dimensions
            GetRoomTransformedBB(TransRoomVertex, this);

            //TaskDialog.Show("Origin", "Width = " + RoomWidth.ToString() + Environment.NewLine
            //    + "Length = " + RoomLength.ToString());



            if (RoomWidth < 14.25 && RoomWidth > 5.5)
            {
                RoomType = "LeftRightSingle";
            }
            if (RoomWidth > 14.25)
            {
                RoomType = "Double";
            }

        }


        static List<XYZ> TransformRoomVertex(List<XYZ> vertexList, Transform tForm)
        {
            List<XYZ> transRoomVertex = new List<XYZ>();

            foreach (XYZ pt in vertexList)
            {
                XYZ rotatedPt = tForm.OfPoint(pt);
                transRoomVertex.Add(rotatedPt);
            }
            return transRoomVertex;
        }
        static void GetRoomTransformedBB(List<XYZ> vertexList, RoomProgramData room1)
        {

            double xMin = (vertexList.OrderBy(p => p.X).FirstOrDefault()).X;
            double xMax = (vertexList.OrderBy(p => p.X).LastOrDefault()).X;
            double yMin = (vertexList.OrderBy(p => p.Y).FirstOrDefault()).Y;
            double yMax = (vertexList.OrderBy(p => p.Y).LastOrDefault()).Y;
            double z = vertexList[0].Z;
            XYZ P = new XYZ(xMin, yMin, z);
            XYZ Q = new XYZ(xMax, yMin, z);
            XYZ R = new XYZ(xMax, yMax, z);
            XYZ S = new XYZ(xMin, yMax, z);

            room1.LeftEdge = new List<XYZ> { P, S };
            room1.RightEdge = new List<XYZ> { Q, R };

            //TaskDialog.Show("Trans Points", " P = " + P.X.ToString() + Environment.NewLine + P.Y.ToString() + Environment.NewLine + P.Z.ToString() + Environment.NewLine
            //    + " Q = " + Q.X.ToString() + Environment.NewLine + Q.Y.ToString() + Environment.NewLine + Q.Z.ToString() + Environment.NewLine
            //    + " R = " + R.X.ToString() + Environment.NewLine + R.Y.ToString() + Environment.NewLine + R.Z.ToString() + Environment.NewLine
            //    + " S = " + S.X.ToString() + Environment.NewLine + S.Y.ToString() + Environment.NewLine + S.Z.ToString() + Environment.NewLine);

            room1.RoomWidth = xMax - xMin;
            room1.RoomLength = yMax - yMin;
            room1.TransformedBB = new List<XYZ> { P, Q, R, S };

        }










        //Get boundary segment array for the room
        static IList<IList<BoundarySegment>> GetBoundaryLoops(Room room)
        {
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);
            return loops;
        }


        //GET ROOM VERTEX
        static List<XYZ> GetRoomVertex(IList<IList<BoundarySegment>> loops)
        {
            //List<XYZ> roomVertices = new List<XYZ>(); //List of all room vertices
            List<XYZ> roomTempVertices = new List<XYZ>();
            List<List<XYZ>> loopVertices = new List<List<XYZ>>();
            List<double> loopArea = new List<double>();
            int index = 0;
            foreach (IList<BoundarySegment> loop in loops)
            {
                roomTempVertices = new List<XYZ>();

                XYZ p0 = null; //previous segment start point
                XYZ p = null; // segment start point
                XYZ q = null; // segment end point

                foreach (BoundarySegment seg in loop)
                {
                    q = seg.GetCurve().GetEndPoint(1);
                    if (p == null)
                    {
                        //roomVertices.Add(seg.GetCurve().GetEndPoint(0));
                        roomTempVertices.Add(seg.GetCurve().GetEndPoint(0));
                        p = seg.GetCurve().GetEndPoint(0);
                        p0 = p;
                        continue;
                    }
                    p = seg.GetCurve().GetEndPoint(0);
                    if (p != null && p0 != null)
                    {
                        if (AreCollinear(p0, p, q, 0.01))//skipping the segments that are collinear
                        {
                            p0 = p;
                            continue;
                        }
                        else
                        {
                            //roomVertices.Add(p);
                            roomTempVertices.Add(p);
                        }
                    }
                    p0 = p;
                }
                loopVertices.Add(roomTempVertices);
                loopArea.Add(GetAreaTrap(loopVertices[index]));
                index++;
            }
            int maxIndex = loopArea.IndexOf(loopArea.Max());
            List<XYZ> roomVertices2 = loopVertices[maxIndex];

            if (roomVertices2.Count() > 3)
            {
                XYZ p0 = roomVertices2[0];
                XYZ p = roomVertices2[1];
                XYZ q = roomVertices2.Last();

                if (AreCollinear(p0, p, q, 0.01))
                {
                    roomVertices2.RemoveAt(0);
                }
            }



            return roomVertices2;
        }

        //Helper Method_Check whether three points are collinear
        static bool AreCollinear(XYZ p1, XYZ p2, XYZ p3, double epsilon)
        {
            bool collinear = false;
            double area = 0.5 * Math.Abs(p1.X * (p2.Y - p3.Y)
                + p2.X * (p3.Y - p1.Y)
                + p3.X * (p1.Y - p2.Y));
            //sometimes area is not exactly zero but is very small number
            if (area < epsilon)
            {
                collinear = true;
            }
            return collinear;
        }

        static double GetAreaTrap(List<XYZ> points)
        {
            points.Add(points[0]);
            double area = Math.Abs(points.Take(points.Count - 1).Select((p, i) => (points[i + 1].X - p.X) * (points[i + 1].Y + p.Y)).Sum() / 2);
            //TaskDialog.Show("Revit Window:", "Area of room = " + area.ToString());
            points.RemoveAt(points.Count - 1);
            return area;
        }


        static void GetRoomDimensions(Room room, RoomProgramData roomPD)
        {
            XYZ minGlobal = roomPD.RoomBB.Min;
            XYZ maxTempGlobal = roomPD.RoomBB.Max;
            XYZ maxGlobal = new XYZ(maxTempGlobal.X, maxTempGlobal.Y, minGlobal.Z);
            XYZ leftMax = new XYZ(minGlobal.X, maxGlobal.Y, minGlobal.Z);
            XYZ rightMin = new XYZ(maxGlobal.X, minGlobal.Y, maxGlobal.Z);

            if (roomPD.RoomOrientation == "horizontal")
            {
                roomPD.LeftEdge = new List<XYZ> { leftMax, maxGlobal };
                roomPD.RightEdge = new List<XYZ> { minGlobal, rightMin };

                roomPD.RoomWidth = Math.Abs(maxGlobal.Y - minGlobal.Y);
                roomPD.RoomLength = Math.Abs(maxGlobal.X - minGlobal.X);
            }
            else
            {
                roomPD.LeftEdge = new List<XYZ> { minGlobal, leftMax };
                roomPD.RightEdge = new List<XYZ> { rightMin, maxGlobal };

                roomPD.RoomWidth = Math.Abs(maxGlobal.X - minGlobal.X);
                roomPD.RoomLength = Math.Abs(maxGlobal.Y - minGlobal.Y);
            }

        }



  




        //MIGHT USE LATER
        //Convex H U L L
        static List<XYZ> GetConvexHull(List<XYZ> points)
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






    }

}
