using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System;

namespace ToolsV2Classes
{
    [Transaction(TransactionMode.Manual)]
    public class GhostedDeskWorkset : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            // Get the Revit application and document 
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                List<FamilyInstance> deskList = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FurnitureSystems)
               .WhereElementIsNotElementType().Cast<FamilyInstance>()
               .Where(x => x.Symbol.Family.Name.Contains("1_Person-Office-Desk"))
               .Where(x => !doc.GetElement(x.LevelId).Name.ToLower().Contains("container")).ToList();

                //List of all the ghosted desk in model
                List<FamilyInstance> ghostedDeskList = new List<FamilyInstance>();

                foreach (FamilyInstance desk in deskList)
                {
                    // Get the "Ghosted" parameter
                    Parameter ghostedParam = desk.LookupParameter("WW-ShowGhosted");

                    // Check if the parameter exists and is a yes/no parameter
                    if (ghostedParam != null)
                    {
                        // Check if the "Ghosted" parameter is set to 1 (On)
                        if (ghostedParam.AsInteger() == 1)
                        {
                            ghostedDeskList.Add(desk);
                        }
                    }
                }
                Workset ghostedWorkset = new FilteredWorksetCollector(doc)
                 .OfKind(WorksetKind.UserWorkset)
                 .FirstOrDefault(workset => workset.Name.Equals("WW-Ghosted Desks"));

                // Start a transaction to modify the worksets
                using (Transaction trans = new Transaction(doc, "Change Desk Workset"))
                {
                    trans.Start();

                    // Iterate through each desk
                    foreach (FamilyInstance ghostedDesk in ghostedDeskList)
                    {

                        ElementId groupId = ghostedDesk.GroupId;
                        if(groupId != null)
                        {
                            Group groupEle = doc.GetElement(groupId) as Group;
                            if (groupEle != null)
                            {
                                if (groupEle.Pinned)
                                {
                                    groupEle.Pinned = false;
                                }
                                groupEle.UngroupMembers();
                            }
                        }
                        Parameter deskWorkset = ghostedDesk.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                        // Change the workset of the desk to "Ghosted workset"
                        if (ghostedDesk.Pinned)
                        {
                            ghostedDesk.Pinned = false;
                        }
                        deskWorkset.Set(ghostedWorkset.Id.IntegerValue);
                    }


                    // Commit the transaction
                    trans.Commit();
                }
                int phyDesk = deskList.Count() - ghostedDeskList.Count();
                TaskDialog.Show("Desk Totals : ", "Count of 1_Person-Office-Desk" + Environment.NewLine +
                    "Total Work Units = " + deskList.Count().ToString()
                    + Environment.NewLine + 
                    "Total Physical Desks = " + phyDesk.ToString()
                    + Environment.NewLine +
                    "Total Ghosted Desks = " + ghostedDeskList.Count().ToString()
                    + Environment.NewLine + Environment.NewLine
                    + "Note: Exec Cabin are not counted in above physical desk count. Only the count of Family instances of 1_Person-Office-Desk is included.");
                
                string toolName = "Ghosted Desk Workset Correction";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Wayfinding Automation";
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
