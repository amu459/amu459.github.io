using System;
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

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class MEPBOQExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                List<Excel.Worksheet> xlWorkSheetList = new List<Excel.Worksheet>();

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks
                    .Open(@"C:\Users\Amrut Modani\Desktop\WeWork-VDC\Learning\WWI_Project Name_MEP Package_BOQ.xlsx");
                
                xlWorkSheetList.Add(xlWorkBook.Sheets[2]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[3]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[5]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[6]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[8]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[10]);
                xlWorkSheetList.Add(xlWorkBook.Sheets[11]);

                #region
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> scheduleCollection = collector.OfClass(typeof(ViewSchedule)).ToElements();
                //HVAC Schedules
                List<ViewSchedule> HVACschedules = new List<ViewSchedule>();
                List<TableData> HVACtables = new List<TableData>();
                List<TableSectionData> HVACsections = new List<TableSectionData>();
                List<int> HVACnRows = new List<int>();

                //EL Schedules
                List<ViewSchedule> ELschedules = new List<ViewSchedule>();
                List<TableData> ELtables = new List<TableData>();
                List<TableSectionData> ELsections = new List<TableSectionData>();
                List<int> ELnRows = new List<int>();

                //FF Schedules
                List<ViewSchedule> FFschedules = new List<ViewSchedule>();
                List<TableData> FFtables = new List<TableData>();
                List<TableSectionData> FFsections = new List<TableSectionData>();
                List<int> FFnRows = new List<int>();

                //FAPA Schedules
                List<ViewSchedule> FAPAschedules = new List<ViewSchedule>();
                List<TableData> FAPAtables = new List<TableData>();
                List<TableSectionData> FAPAsections = new List<TableSectionData>();
                List<int> FAPAnRows = new List<int>();

                ViewSchedule tempSch;
                TableData tempTableData;
                TableSectionData tempTableSectionData;

                foreach (Element e in scheduleCollection)
                {
                    tempSch = e as ViewSchedule;
                    if (tempSch.Name.Contains("ME - "))
                    {
                        tempTableData = tempSch.GetTableData();
                        tempTableSectionData = tempTableData.GetSectionData(SectionType.Body);
                        HVACschedules.Add(tempSch);
                        HVACtables.Add(tempTableData);
                        HVACsections.Add(tempTableSectionData);
                        HVACnRows.Add(tempTableSectionData.NumberOfRows);
                    }
                    if (tempSch.Name.Contains("EL - "))
                    {
                        tempTableData = tempSch.GetTableData();
                        tempTableSectionData = tempTableData.GetSectionData(SectionType.Body);
                        ELschedules.Add(tempSch);
                        ELtables.Add(tempTableData);
                        ELsections.Add(tempTableSectionData);
                        ELnRows.Add(tempTableSectionData.NumberOfRows);
                    }
                    if (tempSch.Name.Contains("FF - "))
                    {
                        tempTableData = tempSch.GetTableData();
                        tempTableSectionData = tempTableData.GetSectionData(SectionType.Body);
                        FFschedules.Add(tempSch);
                        FFtables.Add(tempTableData);
                        FFsections.Add(tempTableSectionData);
                        FFnRows.Add(tempTableSectionData.NumberOfRows);
                    }
                    if (tempSch.Name.Contains("FAPA - "))
                    {
                        tempTableData = tempSch.GetTableData();
                        tempTableSectionData = tempTableData.GetSectionData(SectionType.Body);
                        FAPAschedules.Add(tempSch);
                        FAPAtables.Add(tempTableData);
                        FAPAsections.Add(tempTableSectionData);
                        FAPAnRows.Add(tempTableSectionData.NumberOfRows);
                    }
                }

                #endregion

                List<string> HVACquant = PushToExcel(xlWorkSheetList[0], HVACschedules, HVACnRows);
                List<string> f75quant = PushToExcel(xlWorkSheetList[1], HVACschedules, HVACnRows);
                List<string> BMSquant = PushToExcel(xlWorkSheetList[2], HVACschedules, HVACnRows);
                List<string> ELquant = PushToExcel(xlWorkSheetList[3], ELschedules, ELnRows);
                List<string> FEquant = PushToExcel(xlWorkSheetList[4], FFschedules, FFnRows);
                List<string> FFquant = PushToExcel(xlWorkSheetList[5], FFschedules, FFnRows);
                List<string> FAPAquant = PushToExcel(xlWorkSheetList[6], FAPAschedules, FAPAnRows);

                int j = 0;
                foreach (String st in HVACquant)
                {
                    xlWorkSheetList[0].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in f75quant)
                {
                    xlWorkSheetList[1].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in BMSquant)
                {
                    xlWorkSheetList[2].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in ELquant)
                {
                    xlWorkSheetList[3].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in FEquant)
                {
                    xlWorkSheetList[4].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in FFquant)
                {
                    xlWorkSheetList[5].Cells[j + 1, 5] = st;
                    j++;
                }
                j = 0;
                foreach (String st in FAPAquant)
                {
                    xlWorkSheetList[6].Cells[j + 1, 5] = st;
                    j++;
                }

                xlWorkBook.SaveAs(@"C:\Users\Amrut Modani\Desktop\WeWork-VDC\Learning\abc.xlsx", Excel.XlFileFormat.xlWorkbookDefault);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                foreach (Excel.Worksheet xlws in xlWorkSheetList)
                {
                    Marshal.ReleaseComObject(xlws);
                }
                //close and release
                xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }

        private List<string> PushToExcel(Excel.Worksheet xlWorkSheet, List<ViewSchedule> XXSchedules, List<int> XXnRows)
        {
            Excel.Range range = xlWorkSheet.UsedRange;
            int rowCount = range.Rows.Count;

            List<string> labelValue = new List<string>();
            List<string> unitValue = new List<string>();

            string cellValueLabel;
            string cellValueUnit;
            //Get List of Label and associated Units
            for (int i = 1; i < rowCount; i++)
            {
                cellValueLabel = Convert.ToString(xlWorkSheet.Cells[i, 2].Value);
                cellValueUnit = Convert.ToString(xlWorkSheet.Cells[i, 4].Value);
                if (cellValueLabel == "")
                {
                    labelValue.Add("NA");
                }
                else
                {
                    labelValue.Add(cellValueLabel);
                }

                if (cellValueUnit == "")
                {
                    unitValue.Add("NA");
                }
                else
                {
                    unitValue.Add(cellValueUnit);
                }
            }
            int lineItems = labelValue.Count;
            List<string> quant = new List<string>(new string[lineItems]);

            int counter = 0;
            int nRows = XXnRows[counter];
            foreach (ViewSchedule vs in XXSchedules)
            {
                nRows = XXnRows[counter];
                for (int i = 0; i < nRows; i++)
                {
                    string schLabel = vs.GetCellText(SectionType.Body, i, 0);
                    int labelCounter = 0;
                    foreach (string label in labelValue)
                    {
                        if (label != "" && label == schLabel)
                        {
                            switch (unitValue[labelCounter])
                            {
                                case "Nos":
                                case "Each":
                                case "Lot":
                                case "Lot.":
                                case "LOT":
                                case "Nos.":
                                case "No":
                                case "No.":
                                case "Set":
                                case "Set.":
                                    quant[labelCounter] = vs.GetCellText(SectionType.Body, i, 1);
                                    break;

                                case "Sqmt":
                                case "Sqm":
                                case "Sqm.":
                                    quant[labelCounter] = vs.GetCellText(SectionType.Body, i, 2);
                                    break;

                                case "Rmt":
                                case "RMT":
                                case "Mtrs":
                                case "Mts":
                                case "Mts.":
                                case "Mtr":
                                    quant[labelCounter] = vs.GetCellText(SectionType.Body, i, 3);
                                    break;

                                default:
                                    quant[labelCounter] = "";
                                    break;
                            }
                        }
                        labelCounter++;
                    }
                }
                counter++;
            }
            return quant;
        }
    }
}
