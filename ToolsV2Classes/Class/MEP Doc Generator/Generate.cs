using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Generate : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;

            try
            {
                //Pick Level
                Reference pickedObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                //Retrive Element ID
                ElementId eleId = pickedObj.ElementId;
                Level level = doc.GetElement(eleId) as Level;
                string levelName = level.Name;

                //Get list of views
                List<View> views = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan))
                    .WhereElementIsNotElementType()
                    .Cast<View>().Where(x => x.Name.ToLower().Contains("x floor"))
                    .ToList();
                //Get list of sheets
                List<ViewSheet> sheets = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>().Where(x => x.Name.ToLower().Contains("x floor"))
                    .ToList();

                //Get Sheet Names and Numbers
                List<string> sheetName = new List<string>();
                List<string> sheetNum = new List<string>();
                List<string> sheetNumPrefix = new List<string>();
                List<string> sheetCategory = new List<string>();
                List<string> sheetSubCategory = new List<string>();
                List<string> sheetSeries = new List<string>();
                string sheetNumSuffix = "xx";

                Guid sheetCategoryGuid = sheets[0].LookupParameter("WW-SheetCategory").GUID;
                Guid sheetSubCategoryGuid = sheets[0].LookupParameter("WW-SheetSubCategory").GUID;
                Guid sheetSeriesGuid = sheets[0].LookupParameter("WW-SheetSeries").GUID;

                string trimSheetName = "";
                int index = 0;
                int stringIndex = 0;
                foreach (ViewSheet vs in sheets)
                {
                    trimSheetName = vs.Name.Remove(0, 7);
                    sheetName.Add(levelName.ToUpper() + trimSheetName);
                    sheetNum.Add(vs.SheetNumber);
                    stringIndex = vs.SheetNumber.IndexOf(".");
                    sheetNumPrefix.Add(vs.SheetNumber.Substring(0, stringIndex + 1));
                    sheetCategory.Add(GetSheetParamVal(vs, sheetCategoryGuid));
                    sheetSubCategory.Add(GetSheetParamVal(vs, sheetSubCategoryGuid));
                    sheetSeries.Add(GetSheetParamVal(vs, sheetSeriesGuid));
                }

                FilteredElementCollector titleblocks = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_TitleBlocks);

                //Get Titleblock Id
                FamilySymbol titleBlock = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .FirstOrDefault(q => q.Name == "Horizontal Notes") as FamilySymbol;
                ElementId titleBlockId = titleBlock.Id;

                //Get View Names
                List<string> viewName = new List<string>();
                string trimViewName = "";
                index = 0;
                foreach (View v in views)
                {
                    trimViewName = v.Name.Remove(0, 7);
                    viewName.Add(levelName + trimViewName);
                }

                //Set up Form
                GetLevelNo form1 = new GetLevelNo(commandData);
                form1.ShowDialog();
                sheetNumSuffix = form1.levelNoInput.ToString();

                using (Transaction trans = new Transaction(doc, "Create Views and Sheets"))
                {
                    trans.Start();

                    if (sheetNumSuffix != "xx" && sheetNumSuffix != "")
                    {
                        //Create View
                        index = 0;
                        foreach (View v in views)
                        {
                            ViewPlan vPlan = ViewPlan.Create(doc, v.GetTypeId(), level.Id);
                            vPlan.Name = viewName[index];
                            index++;
                        }

                        //Create Sheets
                        index = 0;
                        foreach (ViewSheet vs in sheets)
                        {
                            titleBlockId = titleBlock.Id;
                            if (vs.SheetNumber.StartsWith("W"))
                            {
                                titleBlockId = titleblocks.Cast<FamilyInstance>().First(q => q.OwnerViewId == vs.Id).GetTypeId();
                            }
                            ViewSheet SHEET = ViewSheet.Create(doc, titleBlockId);
                            SHEET.Name = sheetName[index];
                            if(sheetNumSuffix == "00")
                            {
                                vs.SheetNumber = sheetNumPrefix[index] + "XX";
                            }
                            SHEET.SheetNumber = sheetNumPrefix[index] + sheetNumSuffix;
                            SHEET.get_Parameter(sheetCategoryGuid).Set(sheetCategory[index]);
                            SHEET.get_Parameter(sheetSubCategoryGuid).Set(sheetSubCategory[index]);
                            SHEET.get_Parameter(sheetSeriesGuid).Set(sheetSeries[index]);
                            index++;
                        }
                    }
                    else
                    {
                        string errorTxt = "Level No is incorrect, retry with correct value";
                        TaskDialog.Show("Revit", errorTxt);
                    }
                    trans.Commit();
                }

                string toolName = "MEP Doc";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "MEP Doc";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }
        }
        public String GetSheetParamVal(ViewSheet vs, Guid userParaGuid)
        {
            String paraValue;
            //Parameter x = vs.get_Parameter(userParaGuid);
            paraValue = vs.get_Parameter(userParaGuid).AsString();
            //switch (x.StorageType)
            //{
            //    case StorageType.String:
            //        paraValue = x.AsString();
            //        break;
            //    case StorageType.Integer:
            //        paraValue = x.AsInteger().ToString();
            //        break;
            //    case StorageType.Double:
            //        //covert the number into Metric
            //        paraValue = x.AsValueString();
            //        break;
            //    default:
            //        paraValue = "";
            //        break;
            //}
            return paraValue;
        }
    }
}
