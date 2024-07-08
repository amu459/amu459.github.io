using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolsV2Classes.Class.Lockers;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class LockerCalc : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;



            try
            {


                //Element collections can be done here
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var deskFamilyInstance = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_FurnitureSystems)
                    .Cast<FamilyInstance>()
                    .Where(x => x.Symbol.Family.Name.Contains("1_Person-Office-Desk")).FirstOrDefault();

                Guid isGhostedGuid = deskFamilyInstance.LookupParameter("WW-ShowGhosted").GUID;


                #region Get Rooms from Selection
                //Collect the rooms from user selection, filter out unrequired elements
                List<Room> roomList = new List<Room>();
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (0 == selectedIds.Count)
                {
                    // If no elements are selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + " You haven't selected any Rooms!");
                    goto skipTool;
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
                            if (testRedundant.Area != 0 && programType.ToLower() == "work")
                            {
                                //Add rooms to roomList
                                roomList.Add(elem as Room);
                            }
                        }
                    }
                    if (0 == roomList.Count)
                    {
                        // If no rooms are selected.
                        TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Rooms!");
                        goto skipTool;
                    }
                }
                #endregion


                #region Get Physical Desks
                Room r1 = roomList[0];
                List<FamilyInstance> deskList = GetDesks(r1, isGhostedGuid);
                int X = deskList.Count();

                if (X < 20)
                {
                    TaskDialog.Show("Stupid Title", "Physical Desk Count = " + X.ToString()
                        + Environment.NewLine + Environment.NewLine + "Use Pedestal instead of Lockers for offices with less than 20 physical desks." + Environment.NewLine +"Please refer to standard document for the locker strategy for more information :)");
                    goto skipTool;
                }
                #endregion


                int A1 = (int)Math.Ceiling((decimal)X / 6); //Number of 3x2 = 6 lockers only

                int B1 = (int)Math.Ceiling((decimal)X / 8); ; //Number of 4x2 = 8 lockers only

                int A2 = 0; //Number of 3x2 = 6 lockers in combination
                int B2 = 0; //Number of 4x2 = 8 lockers in combination



                List<List<int>> result = FindOptimalLockers(X);
                int solutionCount = 1;
                string prompt = "Possible Optimal Permutations and Combinations : " + Environment.NewLine;
                
                foreach(List<int> templist in result)
                {
                    prompt += " - " + solutionCount.ToString() + " - "+ Environment.NewLine + "3x2 Locker Units = " + templist[0].ToString() + Environment.NewLine + "4x2 Locker Units = " + templist[1].ToString() 
                        + Environment.NewLine + "Total number of lockers = " + (templist[0] *6 + templist[1]*8).ToString()
                        + Environment.NewLine + Environment.NewLine;
                    solutionCount++;
                }


                //TaskDialog.Show("Revit", "Physical Desk Count = " + X.ToString()
                //        + Environment.NewLine + "Only 3x2 = 6 Lockers = " + A1.ToString()
                //    + Environment.NewLine + "Only 4x2 = 8 Lockers = " + B1.ToString()
                //    + Environment.NewLine + Environment.NewLine + "Combination : " + Environment.NewLine
                //    + "3x2 (6 lockers) = " + A2.ToString() + Environment.NewLine + "4x2 (8 lockers) = " + B2.ToString());


                TaskDialog.Show("Revit", "Physical Desk Count = " + X.ToString() + Environment.NewLine
                        + Environment.NewLine + "Only want to use 3x2 Locker units = " + A1.ToString()
                        + Environment.NewLine + "Total number of lockers = " + (A1*6).ToString()
                        + Environment.NewLine
                        + Environment.NewLine + "Only want to use 4x2 Locker units = " + B1.ToString()
                        + Environment.NewLine + "Total number of lockers = " + (B1 * 8).ToString()

                        + Environment.NewLine + Environment.NewLine + prompt);

                //LockerOutput outputWindow = new LockerOutput();
                //outputWindow.label_Count_6Locs.Content = A1.ToString();
                //outputWindow.label_Count_8Locs.Content = B1.ToString();

                //outputWindow.label_TotalCount_6Locs.Content = (A1 *6).ToString();
                //outputWindow.label_TotalCount_8Locs.Content = (B1 *8).ToString();

                //outputWindow.ShowDialog();



            //using (Transaction tx = new Transaction(doc, "EditedByAPI"))
            //{
            //    tx.Start();
            //    tx.Commit();
            //}


            skipTool:
                string toolName = "Locker Calculator";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Locker Calculator";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }



        //Any Methods Go HERE
        static List<FamilyInstance> GetDesks(Room room, Guid isGhostedGuid)
        {
            /*             
            Returns list of 1_person-office-desk family instances from selected room.
            Ghosted Desks will be ignored. Hot desks are not considered yet.
            */
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter filter
              = new BoundingBoxIntersectsFilter(outline);
            Document doc = room.Document;
            FilteredElementCollector familyInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_FurnitureSystems)
                .WherePasses(filter);
            int roomid = room.Id.IntegerValue;
            List<FamilyInstance> a = new List<FamilyInstance>();
            foreach (FamilyInstance fi in familyInstances)
            {
                if (null != fi.Room
                    && fi.Room.Id.IntegerValue.Equals(roomid)
                    && fi.Symbol.FamilyName.ToLower().Contains("1_person-office-desk")
                    && fi.get_Parameter(isGhostedGuid).AsValueString().Equals("No"))
                {
                    a.Add(fi);
                }
            }
            return a;
        }

        static List<List<int>> FindOptimalLockers(int X)
        {
            int minLockers = int.MaxValue;
            int bestX = 0, bestY = 0;
            List<int> tempminima = new List<int>();
            tempminima.Add(minLockers);
            List<List<int>> comboList = new List<List<int>>();
            comboList.Add(tempminima);
            List<int> tempList = new List<int>();
            for (int x = 1; x <= X; x++)
            {
                for (int y = 1; y <= X; y++)
                {
                    if (6 * x + 8 * y >= X && 6*x + 8*y <= minLockers)
                    {
                        bestX = x;
                        bestY = y;
                        minLockers = 6 * x + 8 * y;
                        if (tempminima.Min() >= minLockers)
                        {
                            tempminima.Add(minLockers);
                            tempList.Add(bestX);
                            tempList.Add(bestY);

                            comboList.Add(tempList);
                            tempList = new List<int>();
                        }
                        minLockers = 6 * x + 8*y;

                    }
                }
            }

            List<List<int>> globalMinamaList = new List<List<int>>();
            int minima = tempminima.Min();
            int count = 0;
            foreach(int min in tempminima)
            {
                if(min == minima)
                {
                    globalMinamaList.Add(comboList[count]);
                }
                count++;
            }


            return globalMinamaList;
        }


    }
}
