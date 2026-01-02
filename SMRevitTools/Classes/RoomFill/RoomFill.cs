using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using SMRevitTools.Classes.RoomFill;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RoomFill : IExternalCommand
    {

        // Implement the Execute method
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            try
            {
                // Create a filtered element collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                List<FamilyInstance> furnitureSystemList = collector.OfCategory(BuiltInCategory.OST_FurnitureSystems).WhereElementIsNotElementType().Cast<FamilyInstance>().Where(x => x.Symbol.FamilyName.ToLower().Contains("roomdata") && !doc.GetElement(x.LevelId).Name.ToLower().Contains("container")).ToList();

                //Family Names
                List<RoomDataFamilyClass> roomDataList = new List<RoomDataFamilyClass>();
                foreach(FamilyInstance fi in furnitureSystemList)
                {
                    Room existingRoom = fi.Room;
                    if (existingRoom == null)
                    {
                        RoomDataFamilyClass roomDataInstance = new RoomDataFamilyClass();
                        roomDataInstance.GetFamilyInstanceData(fi, doc);
                        roomDataList.Add(roomDataInstance);
                    }
                }
                //Workstation List
                List<FamilyInstance> workstationList = collector.OfCategory(BuiltInCategory.OST_FurnitureSystems).WhereElementIsNotElementType().Cast<FamilyInstance>().Where(x => x.Symbol.FamilyName.ToLower().Contains("furn-wk-") && !doc.GetElement(x.LevelId).Name.ToLower().Contains("container")).ToList();
                List<RoomDataFamilyClass> workstationDataList = new List<RoomDataFamilyClass>();
                foreach(FamilyInstance fi in workstationList)
                {
                    Room existingRoom = fi.Room;
                    if (existingRoom == null)
                    {
                        RoomDataFamilyClass workstationDataInstance = new RoomDataFamilyClass();
                        workstationDataInstance.GetFamilyInstanceData(fi, doc);
                        workstationDataList.Add(workstationDataInstance);
                    }
                }








                // Create rooms to specific family locations
                using (Transaction transaction = new Transaction(doc, "Room Fill"))
                {
                    transaction.Start();
                    foreach (RoomDataFamilyClass rdInst in roomDataList)
                    {
                        try
                        {
                            Room room = doc.Create.NewRoom(rdInst.RoomDataLevel, rdInst.CenterLocationUV);
                            if(room.Area > 0 && room.Area<800)
                            {
                                room.LookupParameter("SM-RoomCategory").Set("MEET");
                                room.Name = rdInst.RoomNameSet;


                            }
                            else
                            {
                                doc.Delete(room.Id);
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Revit", "Some error" + ex.Message);
                        }
                    }

                    foreach (RoomDataFamilyClass wkInst in workstationDataList)
                    {
                        try
                        {
                            Room room = doc.Create.NewRoom(wkInst.RoomDataLevel, wkInst.CenterLocationUV);
                            if (room.Area > 0 && room.Area < 800)
                            {
                                room.LookupParameter("SM-RoomCategory").Set("WORK");
                                room.Name = "Office";

                            }
                            else
                            {
                                doc.Delete(room.Id);
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Revit", "Some error : " + ex.Message);
                        }
                    }
                    transaction.Commit();
                }

                // Return success result

                string toolName = "Room Fill Tool";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
            skipped:
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Room Fill Tool";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }
    }
}
