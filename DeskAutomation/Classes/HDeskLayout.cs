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

    public class HDeskLayout : IExternalCommand
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
                FamilySymbol deskType = HelperMethods.GetDeskSymbol(doc);
                if (deskType == null)
                {
                    TaskDialog.Show("Revit Window: ", "Desk Family not found."
                        + Environment.NewLine + "Press F to pay respects."
                        + Environment.NewLine + "(Don't actually press F).");
                    goto cleanup;
                }
                //var doorInfo = HelperMethods.GetDoors(doc);
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
                    roomPD.GetRoomInfo(doc, room, "horizontal");


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

                        List<int> leftDeskValidation0 = CoolerHelperMethods.GetDeskValidation(leftDeskPoints0, roomPD, "left");
                        List<int> rightDeskValidation0 = CoolerHelperMethods.GetDeskValidation(rightDeskPoints0, roomPD, "right");

                        int deskCount0 = CoolerHelperMethods.GetDeskNumbers(leftDeskValidation0);
                        deskCount0 += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation0);


                        //With offset 1
                        List<XYZ> leftDeskPoints1 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPoints1 = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);

                        List<int> leftDeskValidation1 = CoolerHelperMethods.GetDeskValidation(leftDeskPoints1, roomPD, "left");
                        List<int> rightDeskValidation1 = CoolerHelperMethods.GetDeskValidation(rightDeskPoints1, roomPD, "right");

                        int deskCount1 = CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1);
                        deskCount1 += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1);


                        List<XYZ> leftDeskPoints = leftDeskPoints0;
                        List<XYZ> rightDeskPoints = rightDeskPoints0;
                        List<int> leftDeskValidation = leftDeskValidation0;
                        List<int> rightDeskValidation = rightDeskValidation0;
                        if (deskCount1 > deskCount0)
                        {
                            leftDeskPoints = leftDeskPoints1;
                            rightDeskPoints = rightDeskPoints1;
                            leftDeskValidation = leftDeskValidation1;
                            rightDeskValidation = rightDeskValidation1;
                        }

                        bool widthAlgo = true;
                        if (roomPD.RoomWidth < 9.4370)
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
                        if ((roomPD.RoomWidth - 4.71850394) % 9.43700787 < 4.7185)
                        {
                            widthAlgo = false;
                        }

                        //With offset = 1 and Left Orientation
                        List<XYZ> leftDeskPoints1L = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "left", offset1);
                        List<int> leftDeskValidation1L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints1L, roomPD, "left");
                        List<int> rightDeskValidation1L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints1L, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle1L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPointsSingle1L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);

                        List<int> leftDeskValidationSingle1L = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle1L, roomPD, "left");
                        List<int> rightDeskValidationSingle1L = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle1L, roomPD, "right");

                        deskCount1L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1L);
                        deskCount1L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1L);
                        deskCount1L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle1L);
                        if (widthAlgo)
                        {
                            deskCount1L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle1L);

                        }



                        //With offset = 1 and Right Orientation
                        List<XYZ> rightDeskPoints1R = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "right", offset1);
                        List<int> leftDeskValidation1R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints1R, roomPD, "left");
                        List<int> rightDeskValidation1R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints1R, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle1R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset1);
                        List<XYZ> rightDeskPointsSingle1R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset1);

                        List<int> leftDeskValidationSingle1R = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle1R, roomPD, "left");
                        List<int> rightDeskValidationSingle1R = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle1R, roomPD, "right");

                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation1R);
                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation1R);

                        deskCount1R += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle1R);
                        if (widthAlgo)
                        {
                            deskCount1R += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle1R);

                        }


                        //With Offset = 0 and Left orientation
                        List<XYZ> leftDeskPoints0L = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "left", offset0);
                        List<int> leftDeskValidation0L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints0L, roomPD, "left");
                        List<int> rightDeskValidation0L = CoolerHelperMethods.GetDeskValidation(leftDeskPoints0L, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle0L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset0);
                        List<XYZ> rightDeskPointsSingle0L = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset0);

                        List<int> leftDeskValidationSingle0L = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle0L, roomPD, "left");
                        List<int> rightDeskValidationSingle0L = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle0L, roomPD, "right");

                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidation0L);
                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidation0L);
                        deskCount0L += CoolerHelperMethods.GetDeskNumbers(leftDeskValidationSingle0L);

                        if (widthAlgo)
                        {
                            deskCount0L += CoolerHelperMethods.GetDeskNumbers(rightDeskValidationSingle0L);

                        }

                        //With offset = 0 and Right Orientation
                        List<XYZ> rightDeskPoints0R = CoolerHelperMethods.GetDoubleDeskPlacementPoint(roomPD, "right", offset0);
                        List<int> leftDeskValidation0R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints0R, roomPD, "left");
                        List<int> rightDeskValidation0R = CoolerHelperMethods.GetDeskValidation(rightDeskPoints0R, roomPD, "right");

                        List<XYZ> leftDeskPointsSingle0R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.LeftEdge, offset0);
                        List<XYZ> rightDeskPointsSingle0R = CoolerHelperMethods.GetDeskPlacementPoint(roomPD.RightEdge, offset0);

                        List<int> leftDeskValidationSingle0R = CoolerHelperMethods.GetDeskValidation(leftDeskPointsSingle0R, roomPD, "left");
                        List<int> rightDeskValidationSingle0R = CoolerHelperMethods.GetDeskValidation(rightDeskPointsSingle0R, roomPD, "right");

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
                                deskPointsSingle = leftDeskPointsSingle1L;
                                deskValidationSingle = leftDeskValidationSingle1L;
                                deskPoints = leftDeskPoints1L;
                                deskValidationLeft = leftDeskValidation1L;
                                deskValidationRight = rightDeskValidation1L;
                                deskPointsSingle2 = rightDeskPointsSingle1L;
                                deskValidationSingle2 = rightDeskValidationSingle1L;
                                break;

                            case "deskCount1R":
                                deskPointsSingle = rightDeskPointsSingle1R;
                                deskValidationSingle = rightDeskValidationSingle1R;
                                deskPoints = rightDeskPoints1R;
                                deskValidationLeft = leftDeskValidation1R;
                                deskValidationRight = rightDeskValidation1R;
                                deskPointsSingle2 = leftDeskPointsSingle1R;
                                deskValidationSingle2 = leftDeskValidationSingle1R;
                                break;

                            case "deskCount0L":
                                deskPointsSingle = leftDeskPointsSingle0L;
                                deskValidationSingle = leftDeskValidationSingle0L;
                                deskPoints = leftDeskPoints0L;
                                deskValidationLeft = leftDeskValidation0L;
                                deskValidationRight = rightDeskValidation0L;
                                deskPointsSingle2 = rightDeskPointsSingle0L;
                                deskValidationSingle2 = rightDeskValidationSingle0L;
                                break;

                            case "deskCount0R":
                                deskPointsSingle = rightDeskPointsSingle0R;
                                deskValidationSingle = rightDeskValidationSingle0R;
                                deskPoints = rightDeskPoints0R;
                                deskValidationLeft = leftDeskValidation0R;
                                deskValidationRight = rightDeskValidation0R;
                                deskPointsSingle2 = leftDeskPointsSingle0R;
                                deskValidationSingle2 = leftDeskValidationSingle0R;
                                break;

                            default:
                                deskPointsSingle = leftDeskPointsSingle1L;
                                deskValidationSingle = leftDeskValidationSingle1L;
                                deskPoints = leftDeskPoints1L;
                                deskValidationLeft = leftDeskValidation1L;
                                deskValidationRight = rightDeskValidation1L;
                                deskPointsSingle2 = rightDeskPointsSingle1L;
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
                string toolName = "H-Layout";
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
                string toolName = "H-Layout";
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
