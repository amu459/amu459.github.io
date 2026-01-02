using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq.Expressions;

namespace SMRevitTools.Classes.MeetCeiling
{
    public class RoomMethods
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
            return roomList;
        }


        public static List<XYZ> GetRoomVertex(Room room)
        {
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opt);
            
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
            List<XYZ> roomVertices3 = new List<XYZ>();

            foreach(XYZ vert in roomVertices2)
            {
                XYZ roundVertex = new XYZ(Math.Round(vert.X, 3, MidpointRounding.AwayFromZero), Math.Round(vert.Y, 3, MidpointRounding.AwayFromZero), Math.Round(vert.Z));
                roomVertices3.Add(roundVertex);
            }

            return roomVertices3;
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

        public static List<XYZ> GetBoundingBox(List<XYZ> roomVertex)
        {
            
            double commonZ = roomVertex.FirstOrDefault().Z;
            double minX = roomVertex.Min(p => p.X);
            double maxX = roomVertex.Max(p => p.X);
            double minY = roomVertex.Min(p => p.Y);
            double maxY = roomVertex.Max(p => p.Y);

            // Define the rectangle using 4 corner points (bottom-left, bottom-right, top-right, top-left)
            return new List<XYZ>
            {
                new XYZ(minX, minY, commonZ), // Bottom-left
                new XYZ(maxX, minY, commonZ), // Bottom-right
                new XYZ(maxX, maxY, commonZ), // Top-right
                new XYZ(minX, maxY, commonZ)  // Top-left
            };

        }


        public static double FindRotationAngle(List<XYZ> sortedPoints)
        {
            var bottomLeft = sortedPoints[0];
            var bottomRight = sortedPoints[1];

            // Compute angle using Atan2
            double angleRadians = Math.Round(Math.Atan2(bottomRight.Y - bottomLeft.Y, bottomRight.X - bottomLeft.X),3);
            double angleDegrees = angleRadians * (180.0 / Math.PI); // Convert to degrees

            return angleRadians;
        }

        public static List<XYZ> TransformList(Transform tForm, List<XYZ> points)
        {
            List<XYZ> result = new List<XYZ>();
            foreach (XYZ point in points)
            {
                XYZ newPoint = new XYZ(Math.Round(tForm.OfPoint(point).X, 2),
                    Math.Round(tForm.OfPoint(point).Y, 2),
                    point.Z);
                result.Add(newPoint);
            }
            return result;

        }


        public static List<XYZ> GetGridRoomVertex(List<XYZ> roomVertices)
        {
            //roomVertices = roomVertices.OrderBy(p => p.Y).ThenByDescending(p => p.X).ToList();

            //double roundAngle = Math.Round(rotationAngle, 3);
            List<XYZ> roomVerticesOffset = new List<XYZ>();
            XYZ P0 = roomVertices[0];
            XYZ P1 = roomVertices[1];
            XYZ P2 = roomVertices[2];
            XYZ P3 = roomVertices[3];
            double xOffset = 0;
            double yOffset = 0;
            //if (rotationAngle == 0 || rotationAngle == Math.PI * 0.5 || rotationAngle == Math.PI)
            //{
            //    double roomLength = (P1.X - P0.X);
            //    double roomWidth = (P3.Y - P0.Y);

            //    int xTiles = (int)Math.Floor((roomLength - 600 / 304.8) / (600 / 304.8));
            //    int yTiles = (int)Math.Floor((roomWidth - 600 / 304.8) / (600 / 304.8));

            //    //TaskDialog.Show("Revit", "XTiles = " + xTiles.ToString() + " & YTiles = " + yTiles.ToString());
            //    xOffset = 0.5 * (roomLength - xTiles * 600 / 304.8);
            //    yOffset = 0.5 * (roomWidth - yTiles * 600 / 304.8);
            //}
            //else
            //{
            //    double roomLength = (P1.Y - P0.Y)/Math.Sin(roundAngle);
            //    double roomWidth = (P3.Y - P0.Y)/Math.Cos(roundAngle);
            //    int xTiles = (int)Math.Floor((roomLength - 600 / 304.8) / (600 / 304.8));
            //    int yTiles = (int)Math.Floor((roomWidth - 600 / 304.8) / (600 / 304.8));

            //    xOffset = 0.5 * (roomLength - xTiles * 600 / 304.8);
            //    yOffset = 0.5 * (roomWidth - yTiles * 600 / 304.8);
            //}

            double roomLength = (P1.X - P0.X);
            double roomWidth = (P3.Y - P0.Y);

            int xTiles = (int)Math.Floor((roomLength - 600 / 304.8) / (600 / 304.8));
            int yTiles = (int)Math.Floor((roomWidth - 600 / 304.8) / (600 / 304.8));

            //TaskDialog.Show("Revit", "XTiles = " + xTiles.ToString() + " & YTiles = " + yTiles.ToString());
            xOffset = 0.5 * (roomLength - xTiles * 600 / 304.8);
            yOffset = 0.5 * (roomWidth - yTiles * 600 / 304.8);

            XYZ newP0 = new XYZ(P0.X + xOffset, P0.Y + yOffset, P0.Z);
            XYZ newP1 = new XYZ(P1.X - xOffset, P1.Y + yOffset, P1.Z);
            XYZ newP2 = new XYZ(P2.X - xOffset, P2.Y - yOffset, P2.Z);
            XYZ newP3 = new XYZ(P3.X + xOffset, P3.Y - yOffset, P3.Z);
            roomVerticesOffset = new List<XYZ>()
            {newP0, newP1, newP2, newP3 };

            return roomVerticesOffset;
        }


        public static List<XYZ> GetInnerOffsetVertex(List<XYZ> gridVertices)
        {
            XYZ P0 = gridVertices[0];
            XYZ P1 = gridVertices[1];
            XYZ P2 = gridVertices[2];
            XYZ P3 = gridVertices[3];

            double roomLength = (P1.X - P0.X);
            double roomWidth = (P3.Y - P0.Y);


            XYZ newP0 = new XYZ(P0.X + 0.492126, P0.Y + 0.492126, P0.Z);
            XYZ newP1 = new XYZ(P1.X - 0.492126, P1.Y + 0.492126, P1.Z);
            XYZ newP2 = new XYZ(P2.X - 0.492126, P2.Y - 0.492126, P2.Z);
            XYZ newP3 = new XYZ(P3.X + 0.492126, P3.Y - 0.492126, P3.Z);
            List<XYZ> innerVertices= new List<XYZ>()
            {newP0, newP1, newP2, newP3 };

            return innerVertices;
        }

        public static List<string> GetRoomNames (List<Room> rooms)
        {
            List<string> roomNames = new List<string>();
            foreach (Room room in rooms)
            {
                roomNames.Add(room.Name);
            }

            return roomNames;
        }




    }
}