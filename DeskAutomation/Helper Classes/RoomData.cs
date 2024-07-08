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
    public class RoomData
    {
        public Room MyRoom { get; set; } //Room
        public Element RoomLevelElem { get; set; } //Room Level as an Element
        public Level RoomLevel { get; set; } //Room Level as a Level
        public IList<IList<BoundarySegment>> Loops { get; set; } //Boundary segment loops
        public XYZ RoomLocationPoint { get; set; } //Room Location Point
        public FamilyInstance RoomDoor { get; set; }//Door
        public ElementId RoomDoorEdgeId { get; set; } //Element Id of Wall with Door in It
        public Transform RoomTransformObj { get; set; } //Transform of the room vertices wrt door edge angle
        public List<XYZ> RoomVertex { get; set; } //Room vertices 
        public List<XYZ> TransRoomVertex { get; set; } //Transformed Room Vertices
        //public List<XYZ> LeftRightEnds { get; set; } //Extreme points of Door Edge
        //public List<XYZ> DemoCrats { get; set; } //Room vertices sorted such that it starts from LEFT edge of door
        //public List<XYZ> Republicans { get; set; } //Room vertices sorted such that it starts from RIGHT edge of door
        public List<XYZ> LeftEdge { get; set; } //Room Vertices of extreme left edge of door
        public List<XYZ> RightEdge { get; set; } //Room Vertices of extreme right edge of door

        public List<double> AngleToLeftRightEdge { get; set; } //angle of the left and right edge wrt Y Axis

        public double RoomDoorAngle { get; set; } //angle of the door which is considered as angle of the room itself

        //public List<XYZ> RoomConvexHull { get; set; } //Room convex hull derived from Room vertices

        public List<double> RoomDims { get; set; } //Room width and length
        public double RoomWidth { get; set; } //Room width
        public double RoomLength { get; set; } //Room Length
        public string RoomType { get; set; } //To decide which algorithm to use

        public List<XYZ> DoorEdgeEndpoints { get; set; } //
        public List<XYZ> TransformedBB { get; set; }

        public void GetRoomInfo(Document doc, Room room, Dictionary<DoorData, List<Room>> doorInfo)
        {
            MyRoom = room;
            RoomLevel = room.Level;
            RoomLevelElem = room.Level as Element;
            Loops = GetBoundaryLoops(room);
            RoomLocationPoint = ((LocationPoint)room.Location).Point;

            //Find floor below
            Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
            if (floorBelow != null)
            {
                RoomLevelElem = floorBelow;
            }

            RoomDoorEdgeId = GetWallWithDoor(room, doorInfo, Loops, doc, this);

            if (RoomDoorEdgeId != null)
            {
                if (DoorEdgeEndpoints != null)
                {
                    XYZ doorVector = (DoorEdgeEndpoints[1] - DoorEdgeEndpoints[0]).Normalize();
                    RoomDoorAngle = XYZ.BasisX.AngleOnPlaneTo(doorVector, XYZ.BasisZ) * 1;
                }
                else
                {
                    RoomDoorAngle = (RoomDoor.Location as LocationPoint).Rotation;
                }


                //vertices of rooms irrespective of internal holes
                RoomVertex = GetRoomVertex(Loops);

                //Transformed room vertices based on door edge angle
                RoomTransformObj = HelperMethods.GetTransformObj(DoorEdgeEndpoints, -1);
                TransRoomVertex = TransformRoomVertex(RoomVertex, RoomTransformObj);

                //get Bounding box of this transformed room to get Maximum room extents
                RoomDims = GetRoomTransformedBB(TransRoomVertex, this);
                RoomWidth = RoomDims[0];
                RoomLength = RoomDims[1];

                //Yet to restructure
                //LeftRightEnds = GetDoorExtremePoints(TransRoomVertex);
                Transform tFormRe = HelperMethods.GetTransformObj(DoorEdgeEndpoints, 1);

                //DemoCrats = GetLeftWing(TransRoomVertex, LeftRightEnds, this);

                //Republicans = DemoCrats.AsEnumerable().Reverse().ToList();

                LeftEdge = GetLeftEdge(TransRoomVertex, this);
                RightEdge = GetRightEdge(TransRoomVertex, this);

                //AngleToLeftRightEdge = GetEdgeAngles(LeftEdge, RightEdge, RoomLocationPoint, tFormRe);

                AngleToLeftRightEdge = new List<double> { RoomDoorAngle + Math.PI * (0.5),
                    RoomDoorAngle + Math.PI * (1.5)};


                //HelperMethods.PrintPoints(LeftEdge, tFormRe); 
                //HelperMethods.PrintPoints(RightEdge, tFormRe);


                if (RoomWidth < 14.15551 && RoomWidth > 5.5)
                {
                    RoomType = "LeftRightSingle";
                }
                if (RoomWidth > 14.15551 && RoomLength < 23.78608)
                {
                    RoomType = "Double";
                }
                if (RoomWidth > 14.15551 && RoomLength > 23.78608)
                {
                    RoomType = "DoubleLarge";
                }
            }

        }






        //Get boundary segment array for the room
        public static IList<IList<BoundarySegment>> GetBoundaryLoops(Room room)
        {
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);
            return loops;
        }

        //GET DOOR END POINTS
        static ElementId GetWallWithDoor
            (Room room, Dictionary<DoorData, List<Room>> doorInfo, IList<IList<BoundarySegment>> loops, Document doc, RoomData rD)
        {
            ElementId wallId = null;
            int room1Id = 0;
            int room2Id = 0;
            int roomId = room.Id.IntegerValue;
            foreach (var doorObject in doorInfo)
            {
                Room room1 = doorObject.Value[0];
                if (null != room1)
                {
                    room1Id = room1.Id.IntegerValue;
                }
                Room room2 = doorObject.Value[1];
                if (null != room2)
                {
                    room2Id = room2.Id.IntegerValue;
                }

                if (room1Id == roomId || room2Id == roomId)
                {
                    FamilyInstance doorFound = doorObject.Key.Door;
                    rD.RoomDoor = doorFound;
                    wallId = doorFound.Host.Id;
                    Wall doorWall = (Wall)doc.GetElement(wallId);
                    if (doorWall.IsStackedWallMember)
                    {
                        wallId = doorWall.StackedWallOwnerId;
                    }
                    break;
                }
            }

            foreach(var loop in loops)
            {
                foreach(var seg in loop)
                {
                    if (null != seg.ElementId)
                    {
                        ElementId edgeId = seg.ElementId;
                        Element edgeElem = doc.GetElement(edgeId);
                        if (edgeElem != null && edgeElem.Category.Name.ToLower().Contains("wall"))
                        {
                            Wall edgeWall = (Wall)edgeElem;
                            if (edgeWall.IsStackedWallMember)
                            {
                                edgeId = edgeWall.StackedWallOwnerId;
                            }
                            if (wallId.IntegerValue == edgeId.IntegerValue)
                            {
                                var doorEdgeVertices = seg.GetCurve().Tessellate().ToList();
                                rD.DoorEdgeEndpoints = doorEdgeVertices;
                            }
                        }
                    }
                }
            }


            return wallId;
        }


        //GET ROOM VERTEX
        public static List<XYZ> GetRoomVertex(IList<IList<BoundarySegment>> loops)
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

            if(roomVertices2.Count() >3)
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




        //TRANSFORM
        //Transform Room vertices to get horizontal door
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

        //Get Room width from Transformed bounding Box based on Door edge
        static List<double> GetRoomTransformedBB(List<XYZ> vertexList, RoomData room1)
        {
            double roomWidth = 1;
            double roomLength = 1;

            double xMin = (vertexList.OrderBy(p => p.X).FirstOrDefault()).X;
            double xMax = (vertexList.OrderBy(p => p.X).LastOrDefault()).X;
            double yMin = (vertexList.OrderBy(p => p.Y).FirstOrDefault()).Y;
            double yMax = (vertexList.OrderBy(p => p.Y).LastOrDefault()).Y;
            double z = vertexList[0].Z;
            XYZ P = new XYZ(xMin, yMin, z);
            XYZ Q = new XYZ(xMax, yMin, z);
            XYZ R = new XYZ(xMax, yMax, z);
            XYZ S = new XYZ(xMin, yMax, z);

            roomWidth = xMax - xMin;
            roomLength = yMax - yMin;
            room1.TransformedBB = new List<XYZ> { P, Q, R, S };
            List<double> roomDim = new List<double> { roomWidth, roomLength };
            return roomDim;

        }




        //Get Left most EDGE
        static List<XYZ> GetLeftEdge (List<XYZ> transVertex, RoomData roomOb)
        {
            List<XYZ> leftEdge = new List<XYZ>();
            transVertex = transVertex.OrderBy(p => p.X).ToList();
            double epsilon = 0.01;
            double xCoOrdinate1 = transVertex[0].X;

            foreach (var pt in transVertex)
            {
                double xCoOrdinate2 = pt.X;
                if (Math.Abs(xCoOrdinate1 - xCoOrdinate2) < epsilon)
                {
                    leftEdge.Add(pt);
                }
            }
            leftEdge = leftEdge.OrderBy(p => p.Y).ToList();

            return leftEdge;
        }

        //Get Right most EDGE
        static List<XYZ> GetRightEdge(List<XYZ> transVertex, RoomData roomOb)
        {
            List<XYZ> rightEdge = new List<XYZ>();
            transVertex = transVertex.OrderBy(p => p.X).ToList();
            transVertex.Reverse();
            double epsilon = 0.01;
            double xCoOrdinate1 = transVertex[0].X;

            foreach (var pt in transVertex)
            {
                double xCoOrdinate2 = pt.X;
                if (Math.Abs(xCoOrdinate1 - xCoOrdinate2) < epsilon)
                {
                    rightEdge.Add(pt);
                }
            }
            rightEdge = rightEdge.OrderBy(p => p.Y).ToList();

            return rightEdge;
        }





        //UNUSED METHODS : MIGHT DELETE LATER

        //FIND LEFT WING vertices


        //FIND LEFTmost and RIGHTmost points
        static List<XYZ> GetDoorExtremePoints(List<XYZ> transVertex)
        {
            List<XYZ> endPts = new List<XYZ>();


            transVertex = transVertex.OrderBy(p => p.Y).ToList();
            double epsilon = 0.01;
            double yOrdinate1 = transVertex[0].Y;

            List<XYZ> bottomPoints = new List<XYZ>();
            foreach (var pt in transVertex)
            {
                double yOrdinate2 = pt.Y;
                if (Math.Abs(yOrdinate1 - yOrdinate2) < epsilon)
                {
                    bottomPoints.Add(pt);
                }

            }

            bottomPoints = bottomPoints.OrderBy(p => p.X).ToList();

            XYZ leftPt = bottomPoints.First();
            XYZ rightPt = bottomPoints.Last();

            endPts.Add(leftPt);
            endPts.Add(rightPt);

            return endPts;
        }

        static List<XYZ> GetLeftWing (List<XYZ> transVertex, List<XYZ> doorEnds, RoomData rD)
        {
            List<XYZ> leftWing = new List<XYZ>();

            XYZ leftEnd = doorEnds[0];
            XYZ rightEnd = doorEnds[1];
            int leftIndex = transVertex.IndexOf(leftEnd);
            int rightIndex = transVertex.IndexOf(rightEnd);

            if (leftEnd.IsAlmostEqualTo(rightEnd, 0.01))
            {
                leftWing = transVertex.Skip(leftIndex).Concat(transVertex.Take(leftIndex)).ToList();
            }
            else
            {

                if (leftIndex == transVertex.Count()-1)
                {
                    transVertex.Reverse();
                    leftIndex = transVertex.IndexOf(leftEnd);
                }

                else if (leftIndex < rightIndex)
                {
                    transVertex.Reverse();
                    leftIndex = transVertex.IndexOf(leftEnd);
                }
                else if (leftIndex > rightIndex)
                {
                }
                leftWing = transVertex.Skip(leftIndex).Concat(transVertex.Take(leftIndex)).ToList();


            }
            //TaskDialog.Show("Revit", "No of Democrats = " + leftWing.Count().ToString());
            return leftWing;
        }

        //GET only vertex on the LEFT Edge
        static List<XYZ> GetEdge (List<XYZ> leftWing, double maxL)
        {
            List<XYZ> leftEdge = new List<XYZ>();

            double tempLength = 0;
            XYZ previousPt = leftWing[0];
            foreach(var pt in leftWing)
            {
                if(Math.Abs(maxL - tempLength) > 1)
                {
                    leftEdge.Add(pt);
                }
                else if(Math.Abs(maxL - tempLength) <= 1)
                {
                    break;
                }

                double delta = Math.Abs(pt.Y - previousPt.Y);
                tempLength += delta;
                previousPt = pt;
            }

            return leftEdge;
        }

        //For Democrats
        static List<XYZ> GetRoomLeftWing(Room room, ElementId wallId, RoomData rD)
        {
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);


            List<XYZ> allVertex = new List<XYZ>();
            List<XYZ> doorSegment = new List<XYZ>();
            int doorIndex = 0;
            int index = 0;
            bool doorFound = false;
            foreach (var loop in loops)
            {
                //TaskDialog.Show("Revit:", "No of Seg = " + loop.Count().ToString());
                foreach (var seg in loop)
                {
                    
                    List<XYZ> segEndsPts = seg.GetCurve().Tessellate().ToList();
                    if (allVertex.Count() == 0)
                    {
                        allVertex.Add(segEndsPts[0]);
                    }
                    bool tempCheck = false;
                    if(!allVertex.Any(p => p.IsAlmostEqualTo(segEndsPts[1], 0.005)))
                    {
                        allVertex.Add(segEndsPts[1]);
                        index++;
                        tempCheck = true;
                    }


                    if (null != seg.ElementId && doorSegment.Count() == 0)
                    {
                        if(seg.ElementId == wallId)
                        {
                            doorSegment = seg.GetCurve().Tessellate().ToList();
                            rD.DoorEdgeEndpoints = doorSegment;
                            if(tempCheck)
                            {
                                doorIndex = index - 1;
                            }
                            else
                            {
                                doorIndex = index;
                            }
                            doorFound = true;
                        }
                    }
                }
            }

            //TaskDialog.Show("Revit Win", "Door Index = " + doorIndex.ToString() + Environment.NewLine
            //    + "Door Ends = " + Environment.NewLine
            //    + " X1 = " + Math.Round(doorSegment[0].X).ToString()
            //    + " Y1 = " + Math.Round(doorSegment[0].Y).ToString()
            //    + Environment.NewLine
            //    + " X2 = " + Math.Round(doorSegment[1].X).ToString()
            //    + " Y2 = " + Math.Round(doorSegment[1].Y).ToString());

            List<XYZ> demoCats = new List<XYZ>();
            if(doorFound)
            {
                for(int i = doorIndex; i >= 0; i--)
                {
                    demoCats.Add(allVertex[i]);
                }
                for (int i = allVertex.Count()-1; i > doorIndex ; i--)
                {
                    demoCats.Add(allVertex[i]);
                }
            }

            string prompt = "Room Vertex count = " + demoCats.Count().ToString() + Environment.NewLine;


            foreach(var x in demoCats)
            {
                prompt += Environment.NewLine + "X = " + Math.Round(x.X).ToString() + "  _&_  Y = " + Math.Round(x.Y).ToString();
            }


            //TaskDialog.Show("Vertex Count", prompt);



            return demoCats;

        }


        //Get Left and Right edge Angles

        static List<double> GetEdgeAngles(List<XYZ> leftEdgeOg, List<XYZ> rightEdgeOg, XYZ roomLocPt, Transform tFormRe)
        {
            List<double> angles = new List<double>();
            List<XYZ> leftEdge = HelperMethods.TransformList(leftEdgeOg, tFormRe);
            List<XYZ> rightEdge = HelperMethods.TransformList(rightEdgeOg, tFormRe);

            if (leftEdge.Count() > 1)
            {
                XYZ leftEdgeNormal = HelperMethods.GetRoomNormal(leftEdge.First(), leftEdge.Last(), roomLocPt);
                double angleToLeftEdge = XYZ.BasisY.AngleOnPlaneTo(leftEdgeNormal, XYZ.BasisZ);

                angles.Add(angleToLeftEdge);
            }
            else
            {
                angles.Add(0);
            }

            if (rightEdge.Count() > 1)
            {
                XYZ rightEdgeNormal = HelperMethods.GetRoomNormal(rightEdge.First(), rightEdge.Last(), roomLocPt);
                double angleToRightEdge = XYZ.BasisY.AngleOnPlaneTo(rightEdgeNormal, XYZ.BasisZ);

                angles.Add(angleToRightEdge);
            }
            else
            {
                angles.Add(0);
            }

            return angles;
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
