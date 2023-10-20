using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using ToolsV2Classes.Class.ConceptBOQ;
using Autodesk.Revit.UI.Selection;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class ConceptQTO : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            //ToolsV2Classes.Class.ConceptBOQ.UserControl1 form5 = new UserControl1(commandData);
            //form5.Show();
            
            ToolsV2Classes.Class.BOQExportImport.BrowseExcel form4 =
                new ToolsV2Classes.Class.BOQExportImport.BrowseExcel(commandData);
            form4.ShowDialog();
            string boqPath = form4.boqFilePath;
            try
            {
                string fileType = "0";
                if (boqPath.Length >1)
                {
                    fileType = boqPath.Substring(boqPath.Length - 5);
                }
                if (boqPath == "failed" || fileType != ".xlsx")
                {
                    try
                    {
                        int z = 0;
                        int failure = 1 / z;
                        return Result.Succeeded;
                    }
                    catch (Exception e)
                    {
                        //form5.Close();
                        message = "Excel file was not selected, boooo"
                            + Environment.NewLine + "Error : " + e;
                        return Result.Failed;
                    }
                }
            }
            catch (Exception e)
            {
                //form5.Close();
                message = e.Message;
                return Result.Failed;
            }


            #region ExcelDataReader

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(boqPath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            #endregion

            try
            {
                int revitIdColIndex = 3;
                int revitIdRowIndex = 3;
                int revitCatColIndex = 4;
                int revitCatRowIndex = 3;
                int unitColIndex = 5;
                int unitRowIndex = 3;
                int qtyColIndex = 6;
                int qtyRowIndex = 3;

                for (int i = 1; i < 10; i++)
                {
                    bool idFound = false;
                    bool catFound = false;
                    bool unitFound = false;
                    bool qtyFound = false;
                    for (int j = 1; j < 10; j++)
                    {
                        if(xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            string val = xlRange.Cells[i, j].Value2.ToString();
                            if (val.ToLower().Contains("revit_id"))
                            {
                                revitIdColIndex = j;
                                revitIdRowIndex = i;
                                idFound = true;
                            }
                            else if (val.ToLower().Contains("revit_cat"))
                            {
                                revitCatColIndex = j;
                                revitCatRowIndex = i;
                                catFound = true;
                            }
                            else if (val.ToLower().Contains("unit"))
                            {
                                unitColIndex = j;
                                unitRowIndex = i;
                                unitFound = true;
                            }
                            else if (val.ToLower().Contains("quantity"))
                            {
                                qtyColIndex = j;
                                qtyRowIndex = i;
                                qtyFound = true;
                            }
                        }

                    }
                    if (idFound && catFound && unitFound && qtyFound)
                    {
                        break;
                    }
                }


                #region Retriving Type Mark Code Block
                //Selection selection = uidoc.Selection;
                //ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                //foreach (ElementId elId in selectedIds)
                //{
                //    Element el = doc.GetElement(elId);
                //    ElementType wallTypeTest = doc.GetElement(el.GetTypeId()) as ElementType;
                //    var orderedParam = wallTypeTest.GetOrderedParameters();
                //    foreach (var param in orderedParam)
                //    {
                //        if (param.Definition.Name.Equals("Type Mark"))
                //        {
                //            string testStr = param.AsString();
                //            TaskDialog.Show("Revit", "Type mark = " + testStr);
                //        }
                //    }
                //}
                #endregion

                List<LineItem> boqItems = new List<LineItem>();
                for (int i = revitIdRowIndex + 1; i < 150; i++)
                {
                    if(xlRange.Cells[i, revitIdColIndex] != null && xlRange.Cells[i, revitIdColIndex].Value2 != null)
                    {
                        string val = xlRange.Cells[i, revitIdColIndex].Value2.ToString();
                        if (val != null && val.Length > 4 && val.Contains("_"))
                        {
                            if (xlRange.Cells[i, revitCatColIndex] != null && xlRange.Cells[i, revitCatColIndex].Value2 != null)
                            {
                                string catVal = xlRange.Cells[i, revitCatColIndex].Value2.ToString();
                                if (xlRange.Cells[i, unitColIndex] != null && xlRange.Cells[i, unitColIndex].Value2 != null)
                                {
                                    string unitVal = xlRange.Cells[i, unitColIndex].Value2.ToString();
                                    LineItem newItem = new LineItem();
                                    boqItems.Add(newItem.CreateLineItem(val, i+1, unitVal, catVal));
                                }
                            }
                        }
                    }
                }



                //Assembly ceiling door floor furniture system specialityEquip wall


                FilteredElementCollector ceilingCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings);
                FilteredElementCollector doorCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors);
                FilteredElementCollector floorCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors);
                FilteredElementCollector furnitureSysCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FurnitureSystems)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType();
                FilteredElementCollector specialityEqCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SpecialityEquipment)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType();
                FilteredElementCollector wallCol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls);



                foreach (LineItem li in boqItems)
                {
                    if(li.CategoryValue.Contains("ceiling"))
                    {
                        li.ItemQty = CeilingQty(ceilingCol, li, doc);
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }
                    else if (li.CategoryValue.Contains("door"))
                    {
                        li.ItemQty = DoorQty(doorCol, li, doc);
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }
                    else if (li.CategoryValue.Contains("floor"))
                    {
                        li.ItemQty = FloorQty(floorCol, li, doc);
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }
                    else if (li.CategoryValue.Contains("furnituresystem"))
                    {
                        if(li.ParameterValue.ToLower().Contains("office-desk"))
                        {
                            li.ItemQty = OfficeDeskQty(furnitureSysCol, li, doc);
                        }
                        else
                        {
                            li.ItemQty = FurnitureSystemQty(furnitureSysCol, li, doc);
                        }
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }
                    else if (li.CategoryValue.Contains("speciality"))
                    {
                        li.ItemQty = FurnitureSystemQty(specialityEqCol, li, doc);
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }
                    else if (li.CategoryValue.Contains("wall") )
                    {
                        li.ItemQty = WallQty(wallCol, li, doc);
                        xlWorksheet.Cells[li.ItemRow, qtyColIndex] = li.ItemQty;
                        xlApp.Visible = false;
                        xlApp.UserControl = false;
                    }

                }



                string newBoqPath = boqPath.Substring(0, boqPath.LastIndexOf("."));
                string timeStamp = GetTimestamp(DateTime.Now);
                newBoqPath += "_"+ timeStamp+"_RevitExport.xlsx";

                xlWorkbook.SaveAs(newBoqPath);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);


                TaskDialog.Show("Successful Export", "Quantities has been exported successfully at :"
                    + Environment.NewLine + Environment.NewLine + newBoqPath);


                // Return success result
                string toolName = "C&I qto";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }

            catch (Exception e)
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                //form5.Close();
                TaskDialog.Show("Failed Export", "Quantities export has been Failed." 
                    + Environment.NewLine + "Condolences.");
                string toolName = "C&I qto";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }
        }



        public string WallQty 
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            double qtyDoub = 0;

            var filteredWalls = elCol.WhereElementIsNotElementType().OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => (doc.GetElement(x.GetTypeId()) as ElementType).Name.ToLower()
                .Contains(li.ParameterValue.ToLower()))
                .Where(x => !(doc.GetElement(x.LevelId) as Level).Name.ToLower().Contains("container"));

            if (li.ItemUnit.ToLower().Contains("sqm"))
            {

                if (li.ParameterValue.ToLower().Contains("storefront"))
                {
                    foreach (Wall el in filteredWalls)
                    {
                        qtyDoub += el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 2;
                    }
                }
                else
                {
                    foreach (Wall el in filteredWalls)
                    {
                        qtyDoub += el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                    }
                }

                //foreach (Wall el in filteredWalls)
                //{
                //    qtyDoub += el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                //}
                qtyDoub = qtyDoub * 0.092903;
            }
            else if (li.ItemUnit.ToLower().Contains("rmt"))
            {
                foreach (Wall el in filteredWalls)
                {
                    qtyDoub += el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                }
                qtyDoub = qtyDoub * 0.3048;
            }
            qty = Math.Round(qtyDoub, 3).ToString();
            return qty;
        }



        public string CeilingQty
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            double qtyDoub = 0;

            var filteredCeilings = elCol.WhereElementIsNotElementType().OfClass(typeof(Ceiling))
                .Cast<Ceiling>()
                .Where(x => (doc.GetElement(x.GetTypeId()) as ElementType).Name.ToLower()
                .Contains(li.ParameterValue.ToLower()))
                .Where(x => !(doc.GetElement(x.LevelId) as Level).Name.ToLower().Contains("container"));


            if (li.ItemUnit.ToLower().Contains("sqm"))
            {
                foreach (Ceiling el in filteredCeilings)
                {
                    qtyDoub += el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                }
                qtyDoub = qtyDoub * 0.092903;
            }
            qty = Math.Round(qtyDoub, 3).ToString();
            return qty;
        }



        public string DoorQty
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            int qtyDoub = 0;

            if(li.RevitParameter.ToLower().Contains("type"))
            {
                var filteredDoors = elCol.WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(x => x.Symbol.Name.ToLower().Contains(li.ParameterValue.ToLower()))
                .Where(x => !(doc.GetElement(x.LevelId) as Level).Name.ToLower().Contains("container"));

                if (li.ItemUnit.ToLower().Contains("nos"))
                {
                    qtyDoub = filteredDoors.Count();
                }
            }
            else if (li.RevitParameter.ToLower().Contains("family"))
            {
                var filteredDoors = elCol.WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(x => x.Symbol.FamilyName.ToLower().Contains(li.ParameterValue.ToLower()))
                .Where(x => !(doc.GetElement(x.LevelId) as Level).Name.ToLower().Contains("container"));

                if (li.ItemUnit.ToLower().Contains("nos"))
                {
                    qtyDoub = filteredDoors.Count();
                }
            }

            qty = qtyDoub.ToString();
            return qty;
        }



        public string FloorQty
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            double qtyDoub = 0;

            var filteredFloors = elCol.WhereElementIsNotElementType().OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => (doc.GetElement(x.GetTypeId()) as ElementType).Name.ToLower()
                .Contains(li.ParameterValue.ToLower()))
                .Where(x => !(doc.GetElement(x.LevelId) as Level).Name.ToLower().Contains("container"));


            if (li.ItemUnit.ToLower().Contains("sqm"))
            {
                foreach (Floor el in filteredFloors)
                {
                    qtyDoub += el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                }
                qtyDoub = qtyDoub * 0.092903;
            }
            qty = Math.Round(qtyDoub, 3).ToString();
            return qty;
        }


        public string FurnitureSystemQty
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            int qtyDoub = 0;


            if(li.RevitParameter.ToLower().Contains("type"))
            {
                List<FamilyInstance> filteredListType = new List<FamilyInstance>();
                foreach (var ele in elCol)
                {
                    FamilyInstance fInst = ele as FamilyInstance;
                    if (fInst != null)
                    {
                        if (!fInst.Symbol.Family.IsInPlace && fInst.Symbol.Name.ToLower().Contains(li.ParameterValue.ToLower()))
                        {
                            filteredListType.Add(fInst);
                        }
                    }
                }
                if (filteredListType.Count > 0)
                {
                    foreach (var instance in filteredListType)
                    {
                        if (doc.GetElement(instance.LevelId) != null)
                        {
                            if (!(doc.GetElement(instance.LevelId) as Level).Name.ToLower().Contains("container"))
                            {
                                qtyDoub++;
                            }
                        }
                    }
                }
            }


            if (li.RevitParameter.ToLower().Contains("family"))
            {
                List<FamilyInstance> filteredListFamily = new List<FamilyInstance>();
                foreach (var ele in elCol)
                {
                    FamilyInstance fInst = ele as FamilyInstance;
                    if (fInst != null)
                    {
                        if (!fInst.Symbol.Family.IsInPlace && fInst.Symbol.FamilyName.ToLower().Contains(li.ParameterValue.ToLower()))
                        {
                            filteredListFamily.Add(fInst);
                        }
                    }
                }


                if (filteredListFamily.Count > 0)
                {
                    foreach (var instance in filteredListFamily)
                    {
                        if (doc.GetElement(instance.LevelId) != null)
                        {
                            if (!(doc.GetElement(instance.LevelId) as Level).Name.ToLower().Contains("container"))
                            {
                                qtyDoub++;
                            }
                        }
                    }
                }
            }


            qty = qtyDoub.ToString();
            return qty;
        }

        public string OfficeDeskQty
            (FilteredElementCollector elCol, LineItem li, Document doc)
        {
            string qty = "0";
            int qtyDoub = 0;

            List<FamilyInstance> filteredListFamily = new List<FamilyInstance>();
            foreach (var ele in elCol)
            {
                FamilyInstance fInst = ele as FamilyInstance;
                if (fInst != null)
                {
                    if (!fInst.Symbol.Family.IsInPlace && fInst.Symbol.FamilyName.ToLower().Contains(li.ParameterValue.ToLower()))
                    {
                        filteredListFamily.Add(fInst);
                    }
                }
            }


            if (filteredListFamily.Count > 0)
            {
                foreach (var instance in filteredListFamily)
                {
                    if (doc.GetElement(instance.LevelId) != null)
                    {
                        if (!(doc.GetElement(instance.LevelId) as Level).Name.ToLower().Contains("container")
                            && instance.get_Parameter(new Guid("afbfd170-9396-4faf-bd9d-6d03aae40976")).AsValueString().Equals("No"))
                        {
                            qtyDoub++;
                        }
                    }
                }
            }


            qty = qtyDoub.ToString();
            return qty;
        }



        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmss");
        }

    }

}