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
using DeskAutomation.Classes;
using System.Security.Cryptography;

namespace DeskAutomation
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class AngleDeskLayout : IExternalCommand
    {

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
                else if(roomList.Count() > 1)
                {
                    TaskDialog.Show("Multiple Rooms Selected", "The selection contains multiple rooms. Please select a single room for this tool. BYE.");
                    goto cleanup;
                }


                AskAngleOfRoom inputWindow = new AskAngleOfRoom();
                inputWindow.ShowDialog();
                string inputAngle = inputWindow.inputAngle;
                if (inputAngle.ToLower().Contains("cancel"))
                {
                    goto cleanup;
                }
                foreach(char num in inputAngle)
                {
                    if(!char.IsDigit(num)) 
                    {
                        if(!num.ToString().Contains(".") )
                        {
                            if(!num.ToString().Contains("-"))
                            {
                                TaskDialog.Show("Invalid Angle", "Angle has a letter in it! No trolling! Try again with correct angle in degrees in decimal digits. BYE.");
                                goto cleanup;
                            }

                        }

                    }
                }
                double angleInDeg = 0;
                bool success = double.TryParse(inputAngle, out angleInDeg);
                //if (!success)
                //{
                //    TaskDialog.Show("Invaild Angel", "Somtehing Wrogn. Cotnact VDC. BYEEE.");
                //}
                //else
                //{
                //    TaskDialog.Show("Vaild Angle", "Angle = " + angleInDeg.ToString());

                //}


                FamilySymbol deskType = HelperMethods.GetDeskSymbol(doc);
                if (deskType == null)
                {
                    TaskDialog.Show("Revit Window: ", "Desk Family not found."
                        + Environment.NewLine + "Press F to pay respects."
                        + Environment.NewLine + "(Don't actually press F). Please load the desk family. Reach out to VDC team. This is weird not having desk family in the project.");
                    goto cleanup;
                }

                foreach (Room room in roomList)
                {
                    #region Check for Program Type:
                    totalArea += room.Area;
                    string programType;
                    Guid paraGuid = room.LookupParameter("WW-ProgramType").GUID;
                    programType = room.get_Parameter(paraGuid).AsString().ToLower();
                    if(programType != "work")
                    {
                        goto cleanup;
                    }
                    #endregion

                    #region Room Data:
                    RoomProgramData roomPD = new RoomProgramData();
                    roomPD.GetAngledRoomInfo(doc, room, inputAngle);

                    //Room Angle
                    Element host = roomPD.RoomLevelElem;
                    Level level = roomPD.RoomLevel;
                    double roomAngle = roomPD.RoomAngle;
                    XYZ roomLocPt = roomPD.RoomLocationPoint;
                    double maxL = roomPD.RoomLength;
                    double maxW = roomPD.RoomWidth;


                    //Transform objects
                    //Transform tForm = Transform.CreateRotationAtPoint(XYZ.BasisZ, roomAngle, roomPD.RoomVertex[0]);
                    Transform tFormRe = Transform.CreateRotationAtPoint(XYZ.BasisZ, -roomAngle, roomPD.RoomVertex[0]);
                    //Desk List
                    List<FamilyInstance> deskList = new List<FamilyInstance>();

                    #endregion

                    if (roomPD.RoomType.Contains("LeftRightSingle"))
                    {
                        int offset0 = 0;
                        int offset1 = 1;

                        //with offset 0
                        List<XYZ> leftDeskPoints0 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset0);
                        List<XYZ> rightDeskPoints0 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset0);
                        //TaskDialog.Show("Desk Count = ", "Left desk count = " + leftDeskPoints0.Count().ToString()
                        //    + Environment.NewLine + "Right Desk Count = " + rightDeskPoints0.Count().ToString());


                        List<XYZ> leftDesksPointReTrans0 = HelperMethods.TransformList(leftDeskPoints0, tFormRe);
                        List<XYZ> rightDesksPointReTrans0 = HelperMethods.TransformList(rightDeskPoints0, tFormRe);


                        List<int> leftDeskValidation0 = CoolerHelperMethods.GetDeskValidation(leftDesksPointReTrans0, roomPD, "left");
                        List<int> rightDeskValidation0 = CoolerHelperMethods.GetDeskValidation(rightDesksPointReTrans0, roomPD, "right");

                        int deskCount0 = CoolerHelperMethods.GetDeskNumbers(leftDeskValidation0);
                        int tempCount = CoolerHelperMethods.GetDeskNumbers(rightDeskValidation0);

                        deskCount0 += tempCount;


                        //With offset 1
                        List<XYZ> leftDeskPoints1 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPoints1 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);

                        List<XYZ> leftDesksPointReTrans1 = HelperMethods.TransformList(leftDeskPoints1, tFormRe);
                        List<XYZ> rightDesksPointReTrans1 = HelperMethods.TransformList(rightDeskPoints1, tFormRe);

                        List<int> leftDeskValidation1 = CoolerHelperMethods.GetDeskValidation(leftDesksPointReTrans1, roomPD, "left");
                        List<int> rightDeskValidation1 = CoolerHelperMethods.GetDeskValidation(rightDesksPointReTrans1, roomPD, "right");

                        int deskCount1 = CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1);
                        deskCount1 += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1);

                        //TaskDialog.Show("Desk COunt = ", "deskCount0 = " + deskCount0.ToString()
                        //    + Environment.NewLine + "deskCount1 = " + deskCount1.ToString());
                        List<XYZ> leftDeskPoints = leftDesksPointReTrans0;
                        List<XYZ> rightDeskPoints = rightDesksPointReTrans0;
                        List<int> leftDeskValidation = leftDeskValidation0;
                        List<int> rightDeskValidation = rightDeskValidation0;
                        if (deskCount1 > deskCount0)
                        {
                            leftDeskPoints = leftDesksPointReTrans1;
                            rightDeskPoints = rightDesksPointReTrans1;
                            leftDeskValidation = leftDeskValidation1;
                            rightDeskValidation = rightDeskValidation1;
                        }

                        bool widthAlgo = true;
                        if (roomPD.RoomWidth < 9.5)
                        {
                            widthAlgo = false;
                        }

                        using (Transaction trans = new Transaction(doc, "Desktomation: " + room.Name))
                        {
                            trans.Start();
                            //Activate the family type if not activated already
                            if (!deskType.IsActive)
                            {
                                deskType.Activate();
                            }

                            deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, leftDeskPoints, leftDeskValidation, roomPD, "left"));
                            if (deskList.Count() < 1)
                            {
                                widthAlgo = true;
                            }

                            if (widthAlgo)
                            {
                                deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, rightDeskPoints, rightDeskValidation, roomPD, "right"));

                            }

                            trans.Commit();
                        }

                    }


                    if (roomPD.RoomType.Contains("Double"))
                    {
                        int offset0 = 0;
                        int offset1 = 1;

                        int deskCount0L = 0;
                        int deskCount0R = 0;
                        int deskCount1L = 0;
                        int deskCount1R = 0;

                        bool widthAlgo = true;
                        if ((roomPD.RoomWidth - 4.75) % 9.5 < 4.75)
                        {
                            widthAlgo = false;
                        }

                        //With offset = 1 and Left Orientation
                        List<XYZ> leftDeskPoints1L = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "left", offset1);
                        List<XYZ> leftDeskPoints1LTrans = HelperMethods.TransformList(leftDeskPoints1L, tFormRe);



                        List<int> leftDeskValidation1L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints1LTrans, roomPD, "left");
                        List<int> rightDeskValidation1L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints1LTrans, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle1L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPointsSingle1L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);
                        List<XYZ> leftDeskPointsSingle1LTrans = HelperMethods.TransformList(leftDeskPointsSingle1L, tFormRe);
                        List<XYZ> rightDeskPointsSingle1LTrans = HelperMethods.TransformList(rightDeskPointsSingle1L, tFormRe);


                        List<int> leftDeskValidationSingle1L = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle1LTrans, roomPD, "left");
                        List<int> rightDeskValidationSingle1L = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle1LTrans, roomPD, "right");

                        deskCount1L = CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1L);
                        deskCount1L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1L);
                        deskCount1L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle1L);
                        if (widthAlgo)
                        {
                            deskCount1L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle1L);

                        }



                        //With offset = 1 and Right Orientation
                        List<XYZ> rightDeskPoints1R = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "right", offset1);
                        List<XYZ> rightDeskPoints1RTrans = HelperMethods.TransformList(rightDeskPoints1R, tFormRe);

                        List<int> leftDeskValidation1R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints1RTrans, roomPD, "left");
                        List<int> rightDeskValidation1R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints1RTrans, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle1R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPointsSingle1R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);
                        List<XYZ> leftDeskPointsSingle1RTrans = HelperMethods.TransformList(leftDeskPointsSingle1R, tFormRe);
                        List<XYZ> rightDeskPointsSingle1RTrans = HelperMethods.TransformList(rightDeskPointsSingle1R, tFormRe);


                        List<int> leftDeskValidationSingle1R = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle1RTrans, roomPD, "left");
                        List<int> rightDeskValidationSingle1R = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle1RTrans, roomPD, "right");

                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1R);
                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1R);

                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle1R);
                        if (widthAlgo)
                        {
                            deskCount1R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle1R);

                        }


                        //With Offset = 0 and Left orientation
                        List<XYZ> leftDeskPoints0L = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "left", offset0);
                        List<XYZ> leftDeskPoints0LTrans = HelperMethods.TransformList(leftDeskPoints0L, tFormRe);

                        List<int> leftDeskValidation0L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints0LTrans, roomPD, "left");
                        List<int> rightDeskValidation0L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints0LTrans, roomPD, "right");


                        List<XYZ> leftDeskPointsSingle0L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset0);
                        List<XYZ> rightDeskPointsSingle0L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset0);
                        List<XYZ> leftDeskPointsSingle0LTrans = HelperMethods.TransformList(leftDeskPointsSingle0L, tFormRe);
                        List<XYZ> rightDeskPointsSingle0LTrans = HelperMethods.TransformList(rightDeskPointsSingle0L, tFormRe);


                        List<int> leftDeskValidationSingle0L = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle0LTrans, roomPD, "left");
                        List<int> rightDeskValidationSingle0L = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle0LTrans, roomPD, "right");

                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation0L);
                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation0L);
                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle0L);

                        if (widthAlgo)
                        {
                            deskCount0L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle0L);

                        }

                        //With offset = 0 and Right Orientation
                        List<XYZ> rightDeskPoints0R = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "right", offset0);
                        List<XYZ> rightDeskPoints0RTrans = HelperMethods.TransformList(leftDeskPoints0L, tFormRe);

                        List<int> leftDeskValidation0R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints0RTrans, roomPD, "left");
                        List<int> rightDeskValidation0R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints0RTrans, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle0R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset0);
                        List<XYZ> rightDeskPointsSingle0R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset0);
                        List<XYZ> leftDeskPointsSingle0RTrans = HelperMethods.TransformList(leftDeskPointsSingle0R, tFormRe);
                        List<XYZ> rightDeskPointsSingle0RTrans = HelperMethods.TransformList(rightDeskPointsSingle0R, tFormRe);


                        List<int> leftDeskValidationSingle0R = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle0RTrans, roomPD, "left");
                        List<int> rightDeskValidationSingle0R = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle0RTrans, roomPD, "right");

                        deskCount0R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation0R);
                        deskCount0R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation0R);
                        deskCount0R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle0R);

                        if (widthAlgo)
                        {
                            deskCount0R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle0R);

                        }

                        var deskCountDict = new Dictionary<string, int>()
                        {
                            { "deskCount0R", deskCount0R },
                            { "deskCount0L", deskCount0L },
                            { "deskCount1R", deskCount1R },
                            { "deskCount1L", deskCount1L }
                        };


                        string macValueKey = deskCountDict.Aggregate((x, y) => x.Value > y.Value ? x : y).Key;


                        //TaskDialog.Show("Revit:", "Winner = " + macValueKey + Environment.NewLine
                        //    + Environment.NewLine + "deskCount0R" + deskCount0R
                        //    + Environment.NewLine + "deskCount0L" + deskCount0L
                        //    + Environment.NewLine + "deskCount1R" + deskCount1R
                        //    + Environment.NewLine + "deskCount1L" + deskCount1L);

                        List<XYZ> deskPointsSingle = new List<XYZ>();
                        List<int> deskValidationSingle = new List<int>();

                        List<XYZ> deskPoints = new List<XYZ>();
                        List<int> deskValidationLeft = new List<int>();
                        List<int> deskValidationRight = new List<int>();

                        List<XYZ> deskPointsSingle2 = new List<XYZ>();
                        List<int> deskValidationSingle2 = new List<int>();
                        switch (macValueKey)
                        {
                            case "deskCount1L":
                                deskPointsSingle = leftDeskPointsSingle1LTrans;
                                deskValidationSingle = leftDeskValidationSingle1L;
                                deskPoints = leftDeskPoints1LTrans;
                                deskValidationLeft = leftDeskValidation1L;
                                deskValidationRight = rightDeskValidation1L;
                                deskPointsSingle2 = rightDeskPointsSingle1LTrans;
                                deskValidationSingle2 = rightDeskValidationSingle1L;
                                break;

                            case "deskCount1R":
                                deskPointsSingle = rightDeskPointsSingle1RTrans;
                                deskValidationSingle = rightDeskValidationSingle1R;
                                deskPoints = rightDeskPoints1RTrans;
                                deskValidationLeft = leftDeskValidation1R;
                                deskValidationRight = rightDeskValidation1R;
                                deskPointsSingle2 = leftDeskPointsSingle1RTrans;
                                deskValidationSingle2 = leftDeskValidationSingle1R;
                                break;

                            case "deskCount0L":
                                deskPointsSingle = leftDeskPointsSingle0LTrans;
                                deskValidationSingle = leftDeskValidationSingle0L;
                                deskPoints = leftDeskPoints0LTrans;
                                deskValidationLeft = leftDeskValidation0L;
                                deskValidationRight = rightDeskValidation0L;
                                deskPointsSingle2 = rightDeskPointsSingle0LTrans;
                                deskValidationSingle2 = rightDeskValidationSingle0L;
                                break;

                            case "deskCount0R":
                                deskPointsSingle = rightDeskPointsSingle0RTrans;
                                deskValidationSingle = rightDeskValidationSingle0R;
                                deskPoints = rightDeskPoints0RTrans;
                                deskValidationLeft = leftDeskValidation0R;
                                deskValidationRight = rightDeskValidation0R;
                                deskPointsSingle2 = leftDeskPointsSingle0RTrans;
                                deskValidationSingle2 = leftDeskValidationSingle0R;
                                break;

                            default:
                                deskPointsSingle = leftDeskPointsSingle1LTrans;
                                deskValidationSingle = leftDeskValidationSingle1L;
                                deskPoints = leftDeskPoints1LTrans;
                                deskValidationLeft = leftDeskValidation1L;
                                deskValidationRight = rightDeskValidation1L;
                                deskPointsSingle2 = rightDeskPointsSingle1LTrans;
                                deskValidationSingle2 = rightDeskValidationSingle1L;
                                break;

                        }






                        using (Transaction trans = new Transaction(doc, "Desktomation: " + room.Name))
                        {
                            trans.Start();
                            //Activate the family type if not activated already
                            if (!deskType.IsActive)
                            {
                                deskType.Activate();
                            }
                            if (macValueKey.Contains("L"))
                            {
                                deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPointsSingle, deskValidationSingle, roomPD, "left"));
                            }
                            else
                            {
                                deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPointsSingle, deskValidationSingle, roomPD, "right"));
                            }

                            deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPoints, deskValidationLeft, roomPD, "left"));
                            deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPoints, deskValidationRight, roomPD, "right"));

                            if (widthAlgo)
                            {
                                if (macValueKey.Contains("L"))
                                {
                                    deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPointsSingle2, deskValidationSingle2, roomPD, "right"));
                                }
                                else
                                {
                                    deskList.AddRange(CoolerHelperMethods.PlaceDeskSimple(doc, deskType, deskPointsSingle2, deskValidationSingle2, roomPD, "left"));
                                }


                            }
                            trans.Commit();
                        }


                    }
                    deskNumbers += deskList.Count();

                }
                cleanup:
                List<string> deskArea = new List<string>();
                deskArea.Add(deskNumbers.ToString());
                deskArea.Add(totalArea.ToString());
                string toolName = "AngleDeskLayout";
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
                string toolName = "AngleDeskLayout";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateCountDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec, deskArea);
                message = e.Message;
                return Result.Failed;
            }

        }


    }


}
