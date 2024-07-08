using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.UI;
using System.Linq;
using System;
using Autodesk.Revit.DB;

namespace BOQStandardCheck
{
    public static class BOQStandardCheckUtilityClass
    {
        public static List<SheetInfoClass> ReadExcel(string FilePath)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<SheetInfoClass> sheetInfoClassList = new List<SheetInfoClass>();
            List<ExcelWorksheet> sheetList = new List<ExcelWorksheet>();
            ExcelPackage excel = new ExcelPackage(new FileInfo(FilePath));

            foreach (var worksheet in excel.Workbook.Worksheets)
            {
                if (worksheet.Name == "BOQ")
                {
                    sheetList.Add(worksheet);
                }
            }

            foreach (ExcelWorksheet sheet in sheetList)
            {
                int endRowNo = (sheet.Dimension.End.Row);
                SheetInfoClass sheetInfoClass = new SheetInfoClass();
                sheetInfoClass.sheetName = sheet.Name;
                List<rowInfoClass> rowInfoClassList = new List<rowInfoClass>();
                for (int i = 1; i <= endRowNo; i++)
                {
                    rowInfoClass rowInfoClass = new rowInfoClass();
                    var bValue = sheet.Cells[i, 2].Value;
                    if (bValue == null)
                        bValue = "";
                    rowInfoClass.labelValue = bValue.ToString();

                    var aValue = sheet.Cells[i, 1].Value;
                    if (aValue == null)
                        aValue = "";
                    rowInfoClass.omniCodeValue = aValue.ToString();

                    var cValue = sheet.Cells[i, 3].Value;
                    if (cValue == null)
                        cValue = "";
                    rowInfoClass.descriptionValue = cValue.ToString();
                    rowInfoClass.rowNumber = i.ToString();
                    var mergedadress = sheet.MergedCells[i, 3];
                    if (mergedadress != null)
                        rowInfoClass.isMerged = true;
                    rowInfoClassList.Add(rowInfoClass);

                }
                sheetInfoClass.rowInfoClassList = rowInfoClassList;
                sheetInfoClassList.Add(sheetInfoClass);
            }
            return sheetInfoClassList;
        }

        public static ErrorData compareExcel(string excelPath1, string excelPath2)
        {

            List<string> errorList = new List<string>();
            List<SheetInfoClass> sheetInfoClassList1 = new List<SheetInfoClass>();
            sheetInfoClassList1 = ReadExcel(excelPath1);
            List<SheetInfoClass> sheetInfoClassList2 = new List<SheetInfoClass>();
            sheetInfoClassList2 = ReadExcel(excelPath2);
            List<CompareReportClass> compareReportClassList = new List<CompareReportClass>();
            if (sheetInfoClassList1.Count != sheetInfoClassList2.Count)

            {

                string errorCode = "Sheet Count Not Matched in Both Excel.";
                errorList.Add(errorCode);
            }

            foreach (SheetInfoClass sheetInfoClass1 in sheetInfoClassList1)
            {
                CompareReportClass compareReportClass = new CompareReportClass();
                List<string> NotFoundList = new List<string>();
                List<string> NewlyAddedList = new List<string>();
                List<string> NotComparedList = new List<string>();
                SheetInfoClass sheetInfoClass2 = sheetInfoClassList2.FirstOrDefault(x => x.sheetName == sheetInfoClass1.sheetName);
                if (sheetInfoClass2.sheetName != null)
                {
                    foreach (rowInfoClass rowInfoClass1 in sheetInfoClass1.rowInfoClassList)
                    {
                        bool levelfound = false;
                        List<rowInfoClass> MatchedRowInfoClassList = new List<rowInfoClass>();
                        MatchedRowInfoClassList = sheetInfoClass2.rowInfoClassList.Where(x => x.labelValue == rowInfoClass1.labelValue
                        && x.labelValue != "").ToList();
                        if (MatchedRowInfoClassList.Count > 1)
                        {
                            levelfound = true;
                            NotComparedList.Add("Row: " + rowInfoClass1.rowNumber);
                            //string errorCode = string.Format("B{0} Value is not unique, can't compare in Sheet: {1}", rowInfoClass1.rowNumber, sheetInfoClass1.sheetName);
                            //errorList.Add(errorCode);
                        }

                        else
                        {
                            foreach (rowInfoClass rowInfoClass2 in sheetInfoClass2.rowInfoClassList)
                            {
                                if ((rowInfoClass1.labelValue != "") && (rowInfoClass1.labelValue == rowInfoClass2.labelValue))
                                {
                                    levelfound = true;
                                    if ((rowInfoClass1.descriptionValue != rowInfoClass2.descriptionValue) && (rowInfoClass1.isMerged == false))
                                    {
                                        NotFoundList.Add("'Decription' Not Matched in 'Row: " + rowInfoClass1.rowNumber + "' in AOR/EOR Provided BOQ");
                                        //string errorCode = string.Format("Cellvalue C{0} not Matched in Sheet: {1}", rowInfoClass1.rowNumber, sheetInfoClass2.sheetName);
                                        //errorList.Add(errorCode);
                                    }

                                    if (rowInfoClass1.omniCodeValue != rowInfoClass2.omniCodeValue)
                                    {
                                        NotFoundList.Add("'OmniCode' Not Matched in 'Row: " + rowInfoClass1.rowNumber + "' in AOR/EOR Provided BOQ");
                                        //string errorCode = string.Format("Cellvalue A{0} not Matched in Sheet: {1}", rowInfoClass1.rowNumber, sheetInfoClass2.sheetName);
                                        //errorList.Add(errorCode);
                                    }


                                }

                            }
                        }

                        if ((!levelfound) && (rowInfoClass1.labelValue != ""))
                        {
                            NotFoundList.Add("'Label' Not Matched in 'Row: " + rowInfoClass1.rowNumber + "' in AOR/EOR Provided BOQ");
                            //string errorCode = string.Format("CellValue B{0} not Found Sheet: {1}", rowInfoClass1.rowNumber, sheetInfoClass2.sheetName);
                            //errorList.Add(errorCode);
                        }
                    }


                    foreach (rowInfoClass rowInfoClass2 in sheetInfoClass2.rowInfoClassList)
                    {
                        bool isFound = false;
                        foreach (rowInfoClass rowInfoClass1 in sheetInfoClass1.rowInfoClassList)
                        {
                            if ((rowInfoClass2.labelValue != "") && (rowInfoClass2.labelValue == rowInfoClass1.labelValue))
                            {
                                isFound = true;
                                continue;
                            }
                        }
                        if ((!isFound) && (rowInfoClass2.labelValue != ""))
                        {
                            NewlyAddedList.Add("'Label' added in 'Row :" + rowInfoClass2.rowNumber + "' in AOR/EOR Provided BOQ"); ;
                            //string errorCode = string.Format("row {0} newly Added in Sheet: {1}", rowInfoClass2.rowNumber, sheetInfoClass2.sheetName);
                            //errorList.Add(errorCode);
                        }


                    }
                }
                else
                {
                    string errorCode = string.Format("Sheet: {0} Not found", sheetInfoClass1.sheetName);
                    errorList.Add(errorCode);
                }
                compareReportClass.NotFoundList = NotFoundList;
                compareReportClass.NotComparedList = NotComparedList;
                compareReportClass.NewlyAddedList = NewlyAddedList;
                compareReportClass.sheetName = sheetInfoClass1.sheetName;
                compareReportClassList.Add(compareReportClass);

            }
            ErrorData errorData = new ErrorData();
            errorData.otherErrorList = errorList;
            errorData.ComparedReportList = compareReportClassList;
            return errorData;
        }
        public static ErrorData showReport(string excelPath1, string excelPath2)
        {
            ErrorData errorData = compareExcel(excelPath1, excelPath2);
            int totalMismatchedSheet = errorData.ComparedReportList.Count;
            int notFoundCount = 0;
            int newAdditionCount = 0;
            int notComparedCount = 0;
            int otherErrorCount = 0;

            for (int i = 0; i < errorData.ComparedReportList.Count; i++)
            {
                notFoundCount = errorData.ComparedReportList[i].NotFoundList.Count;
                newAdditionCount = errorData.ComparedReportList[i].NewlyAddedList.Count;
                notComparedCount = errorData.ComparedReportList[i].NotComparedList.Count;



            }
            otherErrorCount = errorData.otherErrorList.Count;
            if (errorData.otherErrorList.Count == 0 && errorData.ComparedReportList.Count == 0)
            {
                string ReportMessage = "AOR/EOR BOQ Matched with Standard BOQ";
                TaskDialog.Show("BOQ Standard Check Report Summary", ReportMessage);
            }
            else
            {
                string ReportMessage = string.Format("Total Mismatched Sheet: {0}\n Total Cell value Not Found in new Excel: {1} \n " +
               "Total Cell value added in new Excel: {2} \n Total Cell value unable to compared (Unique Value) : {3} \n" +
               "Other Errors Found :{4}", totalMismatchedSheet.ToString(), notFoundCount.ToString(), newAdditionCount.ToString(), notComparedCount.ToString(), otherErrorCount.ToString());
                TaskDialog.Show("BOQ Standard Check Report Summary", ReportMessage);
            }
            return errorData;

        }
        public static void writeCompareReport(ErrorData errorData, string outputExcelPath)
        {
            int endRowCount = 1;
            // ErrorData errorData = compareExcel(excelPath1, excelPath2);
            if (errorData.otherErrorList.Count == 0 && errorData.ComparedReportList.Count == 0)
            {
                errorData.otherErrorList.Add("AOR/EOR BOQ Matched with Standard BOQ");
            }

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.DefaultRowHeight = 12;
            workSheet.Cells[1, 1].Value = "Sheet Name";
            workSheet.Cells[1, 2].Value = "Not Matched";
            workSheet.Cells[1, 3].Value = "Newly Added";
            workSheet.Cells[1, 4].Value = "Not Compared as Unique Value";
            workSheet.Cells[1, 6].Value = "Other Remarks";

            for (int i = 0; i < errorData.ComparedReportList.Count; i++)
            {
                for (int j = 0; j < errorData.ComparedReportList[i].NotFoundList.Count; j++)
                {
                    workSheet.Cells[j + 2, 1].Value = errorData.ComparedReportList[i].sheetName;
                    workSheet.Cells[j + 2, 2].Value = errorData.ComparedReportList[i].NotFoundList[j];

                }
                endRowCount = Math.Max(endRowCount, errorData.ComparedReportList[i].NotFoundList.Count);
                for (int j = 0; j < errorData.ComparedReportList[i].NewlyAddedList.Count; j++)
                {
                    workSheet.Cells[j + 2, 1].Value = errorData.ComparedReportList[i].sheetName;
                    workSheet.Cells[j + 2, 3].Value = errorData.ComparedReportList[i].NewlyAddedList[j];
                }
                endRowCount = Math.Max(endRowCount, errorData.ComparedReportList[i].NewlyAddedList.Count);
                for (int j = 0; j < errorData.ComparedReportList[i].NotComparedList.Count; j++)
                {
                    workSheet.Cells[j + 2, 1].Value = errorData.ComparedReportList[i].sheetName;
                    workSheet.Cells[j + 2, 4].Value = errorData.ComparedReportList[i].NotComparedList[j];
                }
                endRowCount = Math.Max(endRowCount, errorData.ComparedReportList[i].NotComparedList.Count);



            }

            for (int i = 0; i < errorData.otherErrorList.Count; i++)
            {
                workSheet.Cells[i + 2, 6].Value = errorData.otherErrorList[i];

            }
            endRowCount = Math.Max(endRowCount, errorData.otherErrorList.Count);

            string endCellIndex = "F" + (endRowCount + 1).ToString();
            using (ExcelRange Rng = workSheet.Cells[1, 1, endRowCount + 1, 6])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            using (ExcelRange Rng = workSheet.Cells[1, 1, 1, 6])
            {
                Rng.Style.Font.Bold = true;
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine);
            }
            workSheet.Cells[string.Format("A1:{0}", endCellIndex)].AutoFitColumns();
            if (!File.Exists(outputExcelPath + "\\Standard BOQ Compare Report.xlsx"))
            {
                Stream stream = File.Create(outputExcelPath + "\\Standard BOQ Compare Report.xlsx");
                excel.SaveAs(stream);
                stream.Close();
                TaskDialog.Show("Status", "Compare Report Generated Successfully.");
            }
            else
            {
                TaskDialog.Show("Error Message", "Same Name File already exists in the path");
            }

        }

    }
}
