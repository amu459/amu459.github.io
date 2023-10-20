using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using DeskAutomation.Helper_Classes;


namespace DeskAutomation
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class HelloRevit : IExternalCommand
    {
        public static List<XYZ> endCapDesksLeft = new List<XYZ>();
        public static List<XYZ> endCapDesksRight = new List<XYZ>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            int deskNumbers = 0;
            double totalArea = 0;
            try
            {
                //Filter the selection from Revit and Get only rooms
                List<Room> roomList = HelperMethods.GetRoomsFromSelection(uidoc);

                Dictionary<Room, double> efficiency = new Dictionary<Room, double>();

                if (roomList.Count() == 0)
                {
                    goto cleanup;
                }
                FamilySymbol deskType = HelperMethods.GetDeskSymbol(doc);
                if(deskType == null)
                {
                    TaskDialog.Show("Revit Window: ", "Desk Family not found. Load the standard desk family."
                        + Environment.NewLine + "Press F to pay respects."
                        + Environment.NewLine + "(Don't actually press F).");
                    goto cleanup;
                }

                var doorInfo = HelperMethods.GetDoors(doc);
                foreach (Room room in roomList)
                {
                    totalArea += room.Area;
                    #region Room Data:
                    RoomData roomOb = new RoomData();
                    roomOb.GetRoomInfo(doc, room, doorInfo);

                    if (roomOb.RoomDoorEdgeId == null)
                    {
                        TaskDialog.Show("Revit Window: ", "Door not found for the room."
                            + Environment.NewLine + "No desks for you! :(");
                        goto cleanup;
                    }
                    Element host = roomOb.RoomLevelElem;
                    Level level = roomOb.RoomLevel;
                    XYZ roomLocPt = roomOb.RoomLocationPoint;
                    double maxL = roomOb.RoomLength;
                    double maxW = roomOb.RoomWidth;


                    //Room angle
                    XYZ doorVector = (roomOb.DoorEdgeEndpoints[1] - roomOb.DoorEdgeEndpoints[0]).Normalize();
                    double angleToEdge = XYZ.BasisX.AngleOnPlaneTo(doorVector, XYZ.BasisZ);
                    //Get transformation objects
                    Transform tForm = HelperMethods.GetTransformObj(roomOb.DoorEdgeEndpoints, -1);
                    Transform tFormRe = HelperMethods.GetTransformObj(roomOb.DoorEdgeEndpoints, 1);
                    //Desk List
                    List<FamilyInstance> deskList = new List<FamilyInstance>();

                    List<XYZ> roomVertex = roomOb.RoomVertex;
                    #endregion

                    if (roomOb.RoomType == "LeftRightSingle")
                    {
                        List<XYZ> leftDesksTrans = HelperMethods.TravelFullEdge(roomOb.LeftEdge);
                        List<XYZ> rightDesksTrans = HelperMethods.TravelFullEdge(roomOb.RightEdge);

                        List<XYZ> leftDesks = HelperMethods.TransformList(leftDesksTrans, tFormRe);
                        List<XYZ> rightDesks = HelperMethods.TransformList(rightDesksTrans, tFormRe);
                        bool leftAlgo = true; //for placing desk on left edge of the room
                        bool rightAlgo = true;//for placing desk on right edge of the room
                        bool widthAlgo = true;//I don't remember exact purpose of this boolean, But it seems important based in code below
                        // Based on code below, width checks whether you can fit desks on both edges of the room based on room width.

                        if (roomOb.LeftEdge.Count() < 2)
                        {
                            leftAlgo = false; //validating left edge
                        }

                        if (roomOb.RightEdge.Count() < 2)
                        {
                            rightAlgo = false; //validating right edge
                        }

                        if(roomOb.RoomWidth <= 9.5)
                        {
                            widthAlgo = false; //validating room width to fit desks on both edge of the room
                        }
                        if(!leftAlgo)
                        {
                            widthAlgo = true;
                        }



                        //Find floor below
                        Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
                        if (floorBelow != null)
                        {
                            host = floorBelow;
                        }

                        using (Transaction trans = new Transaction(doc, "Desktomation: "+ room.Name))
                        {
                            trans.Start();
                            //Activate the family type if not activated already
                            if (!deskType.IsActive)
                            {
                                deskType.Activate();
                            }

                            if(leftAlgo)
                            {
                                double angleToLeftEdge = roomOb.AngleToLeftRightEdge[0];

                                deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, leftDesks, host, level, angleToLeftEdge));
                                if(deskList.Count() <1)
                                {
                                    widthAlgo = true;
                                }
                            }

                            if (widthAlgo && rightAlgo)
                            {
                                
                                double angleToRightEdge = roomOb.AngleToLeftRightEdge[1];

                                deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, rightDesks, host, level, angleToRightEdge));
                            }

                            trans.Commit();
                        }
                        double eff = room.Area / (deskList.Count());
                        efficiency.Add(room, eff);

                    }


                    else if(roomOb.RoomType.Contains("Double"))
                    {
                        //XYZ doorMidPoint = ((LocationPoint)roomOb.RoomDoor.Location).Point;
                        //XYZ doorTransPt = tForm.OfPoint(doorMidPoint);
                        //XYZ doorTransLeft = new XYZ(doorTransPt.X - 1.75, doorTransPt.Y, doorTransPt.Z);
                        //XYZ doorTransRight = new XYZ(doorTransPt.X + 1.75, doorTransPt.Y, doorTransPt.Z);
                        endCapDesksLeft = new List<XYZ>();
                        endCapDesksRight = new List<XYZ>();

                        double deskDepth = 4.75;
                        double mRemainder = (maxW - 4.75 + 0.0131234) % (deskDepth * 2); //4mm = 0.0131234 feet for tolerance 
                        //IS 4mm Tolerance even required here? I don't think so. This creates a bug in specific width of room within 4mm window for every 9.5 feet and a single row of desk on the edge gets ignored and desks are not placed

                        //new line to be replaced by old one and to be tested for additional bugs.
                        //double mRemainder = (maxW - 4.75) % (deskDepth * 2);

                        bool rightEdgeAlgo = false;
                        if (mRemainder >= 4.75)
                        {
                            rightEdgeAlgo = true;
                        }


                        string layoutOrientation = "left";


                        //Desk Placement Points Lists
                        List<List<XYZ>> doubleDesksTransLeft = HelperMethods.TravelDoubleBanks3("left", roomOb);
                        List<List<XYZ>> doubleDesksTransRight = HelperMethods.TravelDoubleBanks3("right", roomOb);
                        List<List<XYZ>> doubleDesksLeft = new List<List<XYZ>>();//DOUBLE DESKS LEFT
                        List<List<XYZ>> doubleDesksRight = new List<List<XYZ>>();//DOUBLE DESKS RIGHT



                        //FIND DOUBLE DESK NUMBERS - LEFT ALGO
                        //Find which desks are lying within the room if we start placing desks facing the left edge of the room
                        //This will give a array of 1 and 0 which denotes whether a desk is possible or not
                        List<List<int>> doubleDeskNumsLeft = new List<List<int>>();

                        //ARRAY of desk posibility (0 or 1) which are facing towards left in a double bank of desk
                        List<List<int>> leftDeskNumbsLeft = new List<List<int>>();//Desk Array
                        //ARRAY of desk posibility (0 or 1) which are facing towards ight in a double bank of desk
                        List<List<int>> rightDeskNumbsLeft = new List<List<int>>();//Desk Array


                        List<List<int>> doubleDeskNumsLeft2 = new List<List<int>>();
                        //LIST of desk posibility (0 or 1) which are facing towards left in a double bank of desk
                        List<int> leftDeskNumbsLeft2 = new List<int>();
                        //LIST of desk posibility (0 or 1) which are facing towards left in a double bank of desk
                        List<int> rightDeskNumbsLeft2 = new List<int>();
                        foreach (List<XYZ> dL in doubleDesksTransLeft)
                        {
                            List<XYZ> ttemp = HelperMethods.TransformList(dL, tFormRe);
                            doubleDesksLeft.Add(ttemp);
                            var tempList = HelperMethods.GetDeskNumbers(ttemp, roomOb);
                            leftDeskNumbsLeft.Add(tempList[0]);
                            rightDeskNumbsLeft.Add(tempList[1]);
                            leftDeskNumbsLeft2.AddRange(tempList[0]);
                            rightDeskNumbsLeft2.AddRange(tempList[1]);
                        }
                        doubleDeskNumsLeft.AddRange(leftDeskNumbsLeft);
                        doubleDeskNumsLeft.AddRange(rightDeskNumbsLeft);

                        doubleDeskNumsLeft2.Add(leftDeskNumbsLeft2);
                        doubleDeskNumsLeft2.Add(rightDeskNumbsLeft2);



                        //FIND DOUBLE DESK NUMBERS - RIGHT ALGO
                        List<List<int>> doubleDeskNumsRight = new List<List<int>>();
                        List<List<int>> leftDeskNumbsRight = new List<List<int>>();//Desk Array
                        List<List<int>> rightDeskNumbsRight = new List<List<int>>();//Desk Array

                        List<List<int>> doubleDeskNumsRight2 = new List<List<int>>();
                        List<int> leftDeskNumbsRight2 = new List<int>();
                        List<int> rightDeskNumbsRight2 = new List<int>();
                        foreach (List<XYZ> dL in doubleDesksTransRight)
                        {
                            List<XYZ> ttemp = HelperMethods.TransformList(dL, tFormRe);
                            doubleDesksRight.Add(ttemp);
                            var tempList = HelperMethods.GetDeskNumbers(ttemp, roomOb);
                            leftDeskNumbsRight.Add(tempList[0]);
                            rightDeskNumbsRight.Add(tempList[1]);
                            leftDeskNumbsRight2.AddRange(tempList[0]);
                            rightDeskNumbsRight2.AddRange(tempList[1]);
                        }
                        doubleDeskNumsRight.AddRange(leftDeskNumbsRight); 
                        doubleDeskNumsRight.AddRange(rightDeskNumbsRight);

                        doubleDeskNumsRight2.Add(leftDeskNumbsRight2);
                        doubleDeskNumsRight2.Add(rightDeskNumbsRight2);



                        //FIND EDGE Desks
                        List<XYZ> leftEdgeDesksTrans = HelperMethods.TravelEdgeForDoubleBank3
                            (roomOb.LeftEdge.First(), roomOb.LeftEdge.Last(), "single");
                        List<XYZ> rightEdgeDesksTrans = HelperMethods.TravelEdgeForDoubleBank3
                            (roomOb.RightEdge.First(), roomOb.RightEdge.Last(), "single");


                        List<XYZ> leftEdgeDesks = HelperMethods.TransformList(leftEdgeDesksTrans, tFormRe);//DESKS left
                        List<XYZ> rightEdgeDesks = HelperMethods.TransformList(rightEdgeDesksTrans, tFormRe);//DESKS right

                        List<int> leftEdgeDesksNum = HelperMethods.GetDeskNumbersSingle(leftEdgeDesks, roomOb,"left");
                        List<int> rightEdgeDesksNum = HelperMethods.GetDeskNumbersSingle(rightEdgeDesks, roomOb, "right");




                        int totalLeftDesks = 0;
                        int totalRightDesks = 0;

                        //LEFT algo desks count
                        foreach (List<int> desknum in doubleDeskNumsLeft2)
                        {
                            totalLeftDesks += HelperMethods.GetTotalDeskNumber(desknum);
                        }
                        totalLeftDesks += HelperMethods.GetTotalDeskNumber(leftEdgeDesksNum);
                        if(rightEdgeAlgo)
                        {
                            totalLeftDesks += HelperMethods.GetTotalDeskNumber(rightEdgeDesksNum);
                        }





                        //RIGHT algo desks count
                        foreach (List<int> desknum in doubleDeskNumsRight2)
                        {
                            totalRightDesks += HelperMethods.GetTotalDeskNumber(desknum);
                        }
                        totalRightDesks += HelperMethods.GetTotalDeskNumber(rightEdgeDesksNum);
                        if (rightEdgeAlgo)
                        {
                            totalRightDesks += HelperMethods.GetTotalDeskNumber(leftEdgeDesksNum);
                        }



                        if(totalLeftDesks < totalRightDesks)
                        {
                            layoutOrientation = "right";
                        }
                        double angleToLeftEdge = roomOb.AngleToLeftRightEdge[0];
                        double angleToRightEdge = roomOb.AngleToLeftRightEdge[1];


                        using (Transaction trans = new Transaction(doc, "Desktomation: " + room.Name))
                        {
                            trans.Start();
                            //Activate the family type if not activated already
                            if (!deskType.IsActive)
                            {
                                deskType.Activate();
                            }
                            if(layoutOrientation == "left")
                            {
                                List<int> emptyDeskRow = HelperMethods.GetEmptyRows(leftDeskNumbsLeft);
                                //TaskDialog.Show("Revit WIndow", "Total Desk Rows = ");
                                deskList.AddRange(HelperMethods.PlaceDesks2
                                    (doc, deskType, doubleDesksLeft, leftDeskNumbsLeft, emptyDeskRow, roomOb, angleToLeftEdge));

                                deskList.AddRange(HelperMethods.PlaceDesksSingle
                                    (doc, deskType, leftEdgeDesks, leftEdgeDesksNum, emptyDeskRow, roomOb, angleToLeftEdge));

                                deskList.AddRange(HelperMethods.PlaceDesks2
                                    (doc, deskType, doubleDesksLeft, rightDeskNumbsLeft, emptyDeskRow, roomOb, angleToRightEdge));

                                if (rightEdgeAlgo)
                                {
                                    deskList.AddRange(HelperMethods.PlaceDesksSingle
                                        (doc, deskType, rightEdgeDesks, rightEdgeDesksNum, emptyDeskRow, roomOb, angleToRightEdge));

                                }
                                


                                if(endCapDesksLeft.Count() > 0)
                                {
                                    foreach (XYZ pt in endCapDesksLeft)
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
                                        new XYZ(pt.X+2, pt.Y-3, pt.Z + 2)};

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
                                            FamilyInstance fi = doc.Create.NewFamilyInstance
                                                (pt, deskType, roomOb.RoomLevelElem, roomOb.RoomLevel,
                                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                            //Rotate desk
                                            XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                                            Line axis = Line.CreateBound(pt, pt2);
                                            ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                                            HelperMethods.ChangeOffsetToZero(fi);

                                            deskList.Add(fi);
                                        }

                                    }
                                }

                            }
                            else
                            {
                                List<int> emptyDeskRow = HelperMethods.GetEmptyRows(rightDeskNumbsRight);

                                deskList.AddRange(HelperMethods.PlaceDesks2
                                    (doc, deskType, doubleDesksRight, rightDeskNumbsRight, emptyDeskRow, roomOb, angleToRightEdge));

                                deskList.AddRange(HelperMethods.PlaceDesksSingle
                                    (doc, deskType, rightEdgeDesks, rightEdgeDesksNum, emptyDeskRow, roomOb, angleToRightEdge));

                                deskList.AddRange(HelperMethods.PlaceDesks2
                                    (doc, deskType, doubleDesksRight, leftDeskNumbsRight, emptyDeskRow, roomOb, angleToLeftEdge));

                                if (rightEdgeAlgo)
                                {
                                    deskList.AddRange(HelperMethods.PlaceDesksSingle
                                        (doc, deskType, leftEdgeDesks, leftEdgeDesksNum, emptyDeskRow, roomOb, angleToLeftEdge));

                                }

                                if (endCapDesksRight.Count() > 0)
                                {
                                    foreach (XYZ pt in endCapDesksRight)
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
                                        new XYZ(pt.X+2, pt.Y-3, pt.Z + 2)};

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

                                        if (pointCheck)
                                        {
                                            FamilyInstance fi = doc.Create.NewFamilyInstance
                                                (pt, deskType, roomOb.RoomLevelElem, roomOb.RoomLevel,
                                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                            //Rotate desk
                                            XYZ pt2 = new XYZ(pt.X, pt.Y, pt.Z + 2);

                                            Line axis = Line.CreateBound(pt, pt2);
                                            ElementTransformUtils.RotateElement(doc, fi.Id, axis, angleToEdge);

                                            HelperMethods.ChangeOffsetToZero(fi);

                                            deskList.Add(fi);
                                        }

                                    }
                                }
                            }

                            trans.Commit();
                        }
                        double eff = room.Area / (deskList.Count());
                        efficiency.Add(room, eff);

                    }

                    #region Unused code : DO NOT DELETE
                    //else if (roomOb.RoomType == "DoubleL")
                    //{
                    //    List<int> rowsNumbers = HelperMethods.GetRowNumbers(maxL);

                    //    List<XYZ> leftDesksTrans = HelperMethods.TravelFullEdgeDouble(roomOb.LeftEdge, rowsNumbers, "single");
                    //    List<XYZ> rightDesksTrans = HelperMethods.TravelFullEdgeDouble(roomOb.RightEdge, rowsNumbers, "single");
                    //    List<XYZ> doubleDesksTrans = HelperMethods.TravelDoubleBanks2("left", roomOb);

                    //    List<XYZ> leftDesks = HelperMethods.TransformList(leftDesksTrans, tFormRe);
                    //    List<XYZ> rightDesks = HelperMethods.TransformList(rightDesksTrans, tFormRe);
                    //    List<XYZ> doubleDesks = HelperMethods.TransformList(doubleDesksTrans, tFormRe);


                    //    //Find floor below
                    //    Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
                    //    if (floorBelow != null)
                    //    {
                    //        host = floorBelow;
                    //    }

                    //    double deskDepth = 4.75;
                    //    double mRemainder = (maxW - 4.75) % (deskDepth * 2);
                    //    bool rightEdgeAlgo = false;
                    //    if (mRemainder >= 4.75)
                    //    {
                    //        rightEdgeAlgo = true;
                    //    }

                    //    using (Transaction trans = new Transaction(doc, "Desktomation: " + room.Name))
                    //    {
                    //        trans.Start();
                    //        //Activate the family type if not activated already
                    //        if (!deskType.IsActive)
                    //        {
                    //            deskType.Activate();
                    //        }

                    //        double angleToLeftEdge = roomOb.AngleToLeftRightEdge[0];
                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, leftDesks, host, level, angleToLeftEdge));

                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, doubleDesks, host, level, angleToLeftEdge));

                    //        double angleToRightEdge = roomOb.AngleToLeftRightEdge[1];
                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, doubleDesks, host, level, angleToRightEdge));

                    //        if (rightEdgeAlgo)
                    //        {
                    //            deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, rightDesks, host, level, angleToRightEdge));
                    //        }


                    //        trans.Commit();
                    //    }
                    //    double eff = room.Area / (deskList.Count());
                    //    efficiency.Add(room, eff);

                    //}


                    //else if (roomOb.RoomType == "DoubleR")
                    //{
                    //    List<int> rowsNumbers = HelperMethods.GetRowNumbers(maxL);

                    //    List<XYZ> leftDesksTrans = HelperMethods.TravelFullEdgeDouble(roomOb.LeftEdge, rowsNumbers, "single");
                    //    List<XYZ> rightDesksTrans = HelperMethods.TravelFullEdgeDouble(roomOb.RightEdge, rowsNumbers, "single");
                    //    List<XYZ> doubleDesksTrans = HelperMethods.TravelDoubleBanks2("right", roomOb);

                    //    List<XYZ> leftDesks = HelperMethods.TransformList(leftDesksTrans, tFormRe);
                    //    List<XYZ> rightDesks = HelperMethods.TransformList(rightDesksTrans, tFormRe);
                    //    List<XYZ> doubleDesks = HelperMethods.TransformList(doubleDesksTrans, tFormRe);


                    //    //Find floor below
                    //    Element floorBelow = HelperMethods.FindFloorBelow(doc, room);
                    //    if (floorBelow != null)
                    //    {
                    //        host = floorBelow;
                    //    }

                    //    double deskDepth = 4.75;
                    //    double mRemainder = (maxW - 4.75) % (deskDepth * 2);
                    //    bool leftEdgeAlgo = false;
                    //    if (mRemainder >= 4.75)
                    //    {
                    //        leftEdgeAlgo = true;
                    //    }

                    //    using (Transaction trans = new Transaction(doc, "Desktomation: " + room.Name))
                    //    {
                    //        trans.Start();
                    //        //Activate the family type if not activated already
                    //        if (!deskType.IsActive)
                    //        {
                    //            deskType.Activate();
                    //        }

                    //        double angleToRightEdge = roomOb.AngleToLeftRightEdge[1];

                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, doubleDesks, host, level, angleToRightEdge));

                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, rightDesks, host, level, angleToRightEdge));


                    //        double angleToLeftEdge = roomOb.AngleToLeftRightEdge[0];

                    //        deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, doubleDesks, host, level, angleToLeftEdge));

                    //        if (leftEdgeAlgo)
                    //        {
                    //            deskList.AddRange(HelperMethods.PlaceDesks(room, doc, deskType, leftDesks, host, level, angleToLeftEdge));
                    //        }


                    //        trans.Commit();
                    //    }
                    //    double eff = room.Area / (deskList.Count());
                    //    efficiency.Add(room, eff);

                    //}

                    #endregion

                    deskNumbers += deskList.Count();

                }


                string effPrompt = "Efficiencies Calculated: " + Environment.NewLine;
                foreach(var tt in efficiency)
                {
                    effPrompt += tt.Key.Name + " => " + Math.Round(tt.Value, 2).ToString() + Environment.NewLine;
                }

                //TaskDialog.Show("Desktomation Window:", effPrompt);
            cleanup:
                List<string> deskArea = new List<string>();
                deskArea.Add(deskNumbers.ToString());
                deskArea.Add(totalArea.ToString());
                string toolName = "Autodesk";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Success", doc, uiApp, detlaMilliSec, deskArea);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                List<string> deskArea = new List<string>();
                deskArea.Add(deskNumbers.ToString());
                deskArea.Add(totalArea.ToString());
                string toolName = "Autodesk";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec, deskArea);
                message = e.Message;
                return Result.Failed;
            }

        }






        //Too afraid to scroll down further

        #region Experimental Methods :spooky:

        ////Get the slopes of all room edges (I forgot why I was doing this)
        ////Remembered: To get X and Y axis of the Room orientation
        ////No Need to do all these shenanigans. Imma use door edge slope as Room rotation
        ////K.I.S.S
        //static List<double> GetRoomSlopes (List<XYZ> convexHull)
        //{
        //    List<double> slope = new List<double>();
        //    List<double> edgeLength = new List<double>();
        //    XYZ firstP = convexHull[0];
        //    bool skip = true;
        //    foreach (var x in convexHull)
        //    {
        //        if (skip)
        //        {
        //            skip = false;
        //            continue;
        //        }
        //        XYZ nextP = x;
        //        double dist = nextP.DistanceTo(firstP);
        //        edgeLength.Add(dist);
        //        double x1 = firstP.X;   double y1 = firstP.Y;
        //        double x2 = nextP.X;    double y2 = nextP.Y;
        //        if (Math.Abs(x2-x1) > 0.01)
        //        {
        //            double m = (y2 - y1) / (x2 - x1);
        //            slope.Add(Math.Atan(m)*180/Math.PI);
        //        }
        //        else
        //        {
        //            slope.Add(90);
        //        }
        //        firstP = x;
        //    }

        //    int maxIndex = edgeLength.IndexOf(edgeLength.Max());
        //    double maxSlope = slope[maxIndex];


        //    return slope;
        //}

        ////Get the boundary element containing a door
        ////TODO: Check for Multiple Doors for a single room
        //static ElementId DoorEdge(List<ElementId> roomEdges, Document doc)
        //{
        //    ElementId doorEdgeId = null;
        //    foreach (ElementId edgeId in roomEdges)
        //    {
        //        Element edgeElem = doc.GetElement(edgeId);
        //        if (edgeElem != null)
        //        {
        //            if (edgeElem.Category.Name.Equals("Walls"))
        //            {
        //                Wall edgeWall = edgeElem as Wall;
        //                IList<ElementId> inserts =
        //                    edgeWall.FindInserts(true, true, true, true);

        //                if (inserts.Count() > 0)
        //                {
        //                    doorEdgeId = edgeId;
        //                    break;

        //                }
        //            }
        //        }

        //    }
        //    return doorEdgeId;
        //}


        ////Get Left and Right edge of the Door edge
        //static List<Curve> LeftRightEdge
        //    (Room room, ElementId doorEdge, IList<IList<BoundarySegment>> loops)
        //{
        //    List<Curve> leftRightEdge = new List<Curve>();
        //    var loop = loops[0];
        //    foreach (var seg in loop)
        //    {
        //        //TO DOOOO DO DODO DO Doo
        //    }
        //    return leftRightEdge;
        //}

        ////Mirror for small rectangular Room
        //static List<XYZ> MirroredCoord(List<XYZ> deskPts, XYZ rightEdgeNormal)
        //{
        //    List<XYZ> mirrorDeskPts = new List<XYZ>();
        //    XYZ antiNormal = rightEdgeNormal.Normalize() * 9.5;
        //    Transform tform_nextRow = Transform.CreateTranslation(antiNormal);

        //    foreach (XYZ pt in deskPts)
        //    {
        //        XYZ tpt = tform_nextRow.OfPoint(pt);
        //        mirrorDeskPts.Add(tpt);
        //    }


        //    return mirrorDeskPts;
        //}
        #endregion
    }


}
