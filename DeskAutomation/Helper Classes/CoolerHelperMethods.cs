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
    public class CoolerHelperMethods
    {

        public static Transform GetTransformObj(XYZ origin, double angle, int orientation)
        {

            Transform tForm = Transform.CreateRotationAtPoint(XYZ.BasisZ, angle, origin);

            return tForm;
        }



        public static List<XYZ> GetDeskPlacementPoint(List<XYZ> endPts, int offset)
        {
            List<XYZ> DeskPts = new List<XYZ>();
            double deskWidth = 3.937008;
            XYZ P = endPts[0];
            XYZ Q = endPts[1];
            int deskLimit = 5;
            double dist = P.DistanceTo(Q);
            XYZ v = Q - P;
            XYZ vN = v.Normalize();
            //number of possible desk
            int n = (int)Math.Floor((dist - 0.164042) / deskWidth);
            double remainder = (dist - 0.164042) % deskWidth;

            for (int i = 1; i <= n; i++)
            {
                double d = (i - 1) * deskWidth + 2.05052 + remainder* offset;
                XYZ tPoint = P + d * vN;

                DeskPts.Add(tPoint);
            }

            return DeskPts;
        }


        public static List<int> GetDeskValidation(List<XYZ> deskPoints, RoomProgramData roomPD, string dir)
        {
            List<int> deskValidation = new List<int>();
            Room room = roomPD.MyRoom;
            double angle = 0;

            if (dir == "right")
            {
                angle = roomPD.RoomRightAngle;
            }
            if (dir == "left")
            {
                angle = roomPD.RoomLeftAngle;
            }
            foreach (XYZ pt in deskPoints)
            {

                List<XYZ> clearancePts = new List<XYZ>
                { new XYZ(pt.X, pt.Y-4.7185+0.00656168, pt.Z + 2),//keeping 2mm internal tolerance
                new XYZ(pt.X+1.9685-0.00656168, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-1.9685+0.00656168, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X-1.9685+0.00656168, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X+1.9685-0.00656168, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-2, pt.Z + 2),
                new XYZ(pt.X, pt.Y-0.00656168, pt.Z + 2),
                new XYZ(pt.X, pt.Y-1, pt.Z+2),
                new XYZ(pt.X, pt.Y-3, pt.Z+2),
                new XYZ(pt.X-1.9685+0.00656168, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X+1.9685-0.00656168, pt.Y-3, pt.Z + 2),
                new XYZ(pt.X-1.9685+0.00656168, pt.Y-4.7185+0.00656168, pt.Z + 2),
                new XYZ(pt.X+1.9685-0.00656168, pt.Y-4.7185+0.00656168, pt.Z + 2)};

                bool pointCheck = true;
                Transform tFormNew = Transform.CreateRotationAtPoint(XYZ.BasisZ, angle, pt);
                foreach (XYZ point in clearancePts)
                {
                    XYZ transPoint = tFormNew.OfPoint(point);

                    if (!room.IsPointInRoom(transPoint))
                    {
                        pointCheck = false;
                        break;
                    }

                }

                //TaskDialog.Show("Revit", prompt);
                if (pointCheck)
                {
                    deskValidation.Add(1);
                }
                else
                {
                    deskValidation.Add(0);
                }

            }

            return deskValidation;
        }



        public static List<FamilyInstance> PlaceDeskSimple
            (Document doc, FamilySymbol deskType, List<XYZ> deskPoints, List<int> deskValidation, RoomProgramData roomPD, string dir)
        {
            List<FamilyInstance> deskPlaced = new List<FamilyInstance>();
            double angle = 0;
            if (dir == "right")
            {
                angle = roomPD.RoomRightAngle;
            }
            if (dir == "left")
            {
                angle = roomPD.RoomLeftAngle;
            }

            int count = 0;
            foreach (XYZ pt in deskPoints)
            {
                if(deskValidation[count] == 0)
                {
                    count++;
                    continue;
                }
                else
                {
                    //Create Family Instance
                    FamilyInstance fi = doc.Create.NewFamilyInstance
                        (pt, deskType, roomPD.RoomLevelElem, roomPD.RoomLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    //Rotate desk
                    XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                    Line axis = Line.CreateBound(pt, pt2);
                    ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);

                    HelperMethods.ChangeOffsetToZero(fi);

                    deskPlaced.Add(fi);
                }

                count++;
            }

            return deskPlaced;
        }




        public static List<XYZ> GetDoubleDeskPlacementPoint
            (RoomProgramData roomPD, string dir, int offset)
        {
            List<XYZ> leftEndPts = roomPD.LeftEdge;
            List<XYZ> rightEndPts = roomPD.RightEdge;

            List<XYZ> deskPts = new List<XYZ>();
            double deskDepth = 4.7185;

            double maxL = roomPD.RoomLength;
            double maxW = roomPD.RoomWidth;

            maxW -= deskDepth;
            XYZ P = leftEndPts[0];
            XYZ Q = leftEndPts[1];

            int direction = 1;
            if(dir =="right")
            {
                direction = -1;
                P = rightEndPts[0];
                Q = rightEndPts[1];
            }
            int m = (int)Math.Floor(maxW / (deskDepth * 2));
            if(roomPD.RoomOrientation == "vertical")
            {
                for (int i = 1; i <= m; i++)
                {
                    double d = P.X + direction * i * deskDepth * 2;

                    XYZ startPt = new XYZ(d, P.Y, P.Z);
                    XYZ endPt = new XYZ(d, Q.Y, Q.Z);
                    List<XYZ> edgeList = new List<XYZ> { startPt, endPt };

                    deskPts.AddRange(GetDeskPlacementPoint(edgeList, offset));
                }
            }
            else if (roomPD.RoomOrientation == "horizontal")
            {
                for (int i = 1; i <= m; i++)
                {
                    double d = Q.Y - direction * i * deskDepth * 2;

                    XYZ startPt = new XYZ(P.X, d, P.Z);
                    XYZ endPt = new XYZ(Q.X, d, Q.Z);
                    List<XYZ> edgeList = new List<XYZ> { startPt, endPt };

                    deskPts.AddRange(GetDeskPlacementPoint(edgeList, offset));
                }
            }


            return deskPts;
        }



        public static int GetDeskNumbers (List<int> DeskValidation)
        {
            int desksNum = 0;

            foreach(int i in DeskValidation)
            {
                if(i==1)
                {
                    desksNum++;
                }
            }

            return desksNum;
        }



    }
}
