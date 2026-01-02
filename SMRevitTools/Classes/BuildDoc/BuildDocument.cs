/*
 * Tool Name = Build Document
 * Description = Creates Views, Schedules, Sheets for the floors selected in the project
 * Requirement = Scope box, Levels to be created as required
 */


using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SMRevitTools.Classes.BuildDoc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class BuildDocument : IExternalCommand
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
                //GET LIST OF LEVELS FROM REVIT
                string[] levelNames = GetLevelNames(doc);
                //Set up Form
                LevelSelect levelSelectWindow = new LevelSelect(uidoc, levelNames);
                levelSelectWindow.ShowDialog();

                List<BD_Floors> floorObjextList = levelSelectWindow.floorObjList;
                List<string> selectedLevelNames = new List<string>();
                foreach (var item in floorObjextList)
                {
                    if (item.Check_Status == true)
                    {
                        selectedLevelNames.Add(item.Floor_Name);
                    }
                }

                if(selectedLevelNames.Count() < 1)
                {
                    TaskDialog.Show("Revit", "No Levels Selected!");
                    goto skipped;
                }

                
                //List of selected levels by user
                List<Level> selectedLevels = new List<Level>();
                List<string> geometricLevels = new List<string>();
                foreach (string levelName in selectedLevelNames)
                {
                    Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name.ToLower().Contains(levelName.ToLower())).FirstOrDefault();
                    string geometricLevel = level.LookupParameter("SM-Geometric Level").AsString();
                    if (geometricLevel == null)
                    {
                        geometricLevels.Add("null");
                        TaskDialog.Show("Revit", "Level: " + levelName + " doesn't have a valid value in SM-Geometric Level paratmer. Please follow process.");
                        goto skipped;
                    }
                    geometricLevels.Add(geometricLevel);
                    selectedLevels.Add(level);
                }
                if (geometricLevels.Count() != geometricLevels.Distinct().Count())
                {
                    TaskDialog.Show("Revit", "SM-Geometric Level paratmer is not unique to all the levels. Please recheck and try again.");
                    goto skipped;
                }



                //Get list of views with XFloor in Model
                List<View> xFloorViews = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan))
                    .WhereElementIsNotElementType()
                    .Cast<View>().Where(x => x.Name.ToLower().Contains("x floor"))
                    .ToList();
                //Get list of sheets with XFloor in Model
                List<ViewSheet> xFloorSheets = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>().Where(x => x.Name.ToLower().Contains("x floor"))
                    .ToList();




                ////Get TitleBlock family
                //FilteredElementCollector titleblocks = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_TitleBlocks);
                //Get Titleblock Id
                FamilySymbol titleBlock = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .FirstOrDefault(q => q.Name == "A1") as FamilySymbol;
                ElementId titleBlockId = titleBlock.Id;

                //Create SMSheetClass objects

                foreach (Level lvl in selectedLevels)
                {
                    List<SMSheetClass> sheetObjectList = new List<SMSheetClass>();
                    List<View> viewsFromSheet = new List<View>();
                    List<View> legendsFromSheet = new List<View>();

                    foreach (ViewSheet vSheet in xFloorSheets)
                    {
                        SMSheetClass smSheetObject = new SMSheetClass();
                        smSheetObject.GetSheetData(vSheet, doc, lvl);
                        sheetObjectList.Add(smSheetObject);
                    }

                    using (Transaction trans = new Transaction(doc, "SM-Build Document"))
                    {
                        trans.Start();
                        //foreach (View v in xFloorViews)
                        //{
                        //    string trimViewName = v.Name.Remove(0, 7);
                        //    string newViewName = lvl.Name + trimViewName;
                        //    ViewPlan vPlan = ViewPlan.Create(doc, v.GetTypeId(), lvl.Id);
                        //    vPlan.Name = newViewName;

                        //}
                        foreach (SMSheetClass vs in sheetObjectList)
                        {
                            titleBlockId = titleBlock.Id;
                            ViewSheet SHEET = ViewSheet.Create(doc, titleBlockId);
                            SHEET.Name = vs.SheetName.ToUpper();
                            SHEET.SheetNumber = vs.SheetNum;
                            SHEET.LookupParameter("SM-SheetCategory").Set(vs.SheetCategory);
                            SHEET.LookupParameter("SM-SheetSubCategory").Set(vs.SheetSubCategory);
                            SHEET.LookupParameter("SM-SheetSeries").Set(vs.SheetSeries);

                            if(vs.ViewsInSheet.Count > 0)
                            {
                                foreach(View v in vs.ViewsInSheet)
                                {
                                    foreach (Viewport vp in vs.ViewportsInSheet)
                                    {
                                        if (vp.ViewId == v.Id)
                                        {
                                            string trimViewName = v.Name.Remove(0, 7);
                                            string newViewName = lvl.Name + trimViewName;
                                            ViewPlan vPlan = ViewPlan.Create(doc, v.GetTypeId(), lvl.Id);
                                            vPlan.Name = newViewName;

                                            XYZ viewportCenter = vp.GetBoxCenter();
                                            Viewport test = Viewport.Create(doc, SHEET.Id, vPlan.Id, viewportCenter);
                                            ElementId vpTypeId = vp.GetTypeId();
                                            test.ChangeTypeId(vpTypeId);
                                        }
                                    }

                                }
                            }

                            if (vs.LegendsInSheet.Count > 0)
                            {
                                foreach (View v in vs.LegendsInSheet)
                                {
                                    foreach(Viewport vp in vs.ViewportsInSheet)
                                    {
                                        if(vp.ViewId == v.Id)
                                        {
                                            XYZ viewportCenter = vp.GetBoxCenter();
                                            Viewport test = Viewport.Create(doc, SHEET.Id, v.Id, viewportCenter);
                                            ElementId vpTypeId = vp.GetTypeId();
                                            test.ChangeTypeId(vpTypeId);
                                        }
                                    }

                                }
                            }


                        }

                        trans.Commit();
                    }

                }

                string toolName = "Build Document Tool";
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
                string toolName = "Build Document Tool";
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
            paraValue = vs.get_Parameter(userParaGuid).AsString();
            return paraValue;
        }


        static string[] GetLevelNames(Document doc)
        {
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc).OfClass(typeof(Level));

            List<Level> levels = levelCollector
                .Cast<Level>()
                .Where(x => !x.Name.ToLower().Contains("container")).ToList();

            int levelCount = levels.Count();
            string[] levelNames = new string[levelCount];
            int count = 0;
            foreach (Level level in levels)
            {
                levelNames.SetValue(level.Name, count);
                count++;
            }

            return levelNames;
        }
    }
}
