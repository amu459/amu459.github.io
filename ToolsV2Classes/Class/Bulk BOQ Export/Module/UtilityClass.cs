using System;
using System.Collections.Generic;
using System.Net.Http;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;


namespace ToolsV2Classes
{
    public static class UtilityClass
    {
        public static List<FamilyCostInfoClass> GetAirtableValue(string AirtableApiKey, string BaseId, string TableName)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + AirtableApiKey);
                string apiUrl = $"https://api.airtable.com/v0/{BaseId}/{TableName}";

                bool hasMoreRecords = true;
                string offset = null;
                List<FamilyCostInfoClass> familyCostInfoClassList = new List<FamilyCostInfoClass>();
                while (hasMoreRecords == true)
                {
                    string requestUrl = apiUrl;
                    if (!string.IsNullOrEmpty(offset))
                    {
                        requestUrl += $"?offset={offset}";
                    }
                    HttpResponseMessage response = client.GetAsync(requestUrl).Result;

                    if (response.IsSuccessStatusCode)
                    {

                        string json = response.Content.ReadAsStringAsync().Result;
                        //JObject data = JObject.Parse(json);
                        // dynamic  data= JsonConvert.DeserializeObject(json);

                        Root rootObject = JsonConvert.DeserializeObject<Root>(json);
                        //hasMoreRecords = !string.IsNullOrEmpty(rootObject.Offset);

                        int totalCount = rootObject.records.Count;

                        for (int i = 0; i < totalCount; i++)
                        {
                            FamilyCostInfoClass familyCostInfoClass = new FamilyCostInfoClass();
                            //string desiredValue = data["records"][i]["fields"][FieldName].ToString(); 
                            familyCostInfoClass.TypeName = rootObject.records[i].fields.SKU.ToString();
                            try
                            {
                                if(rootObject.records[i].fields.UnitPrice!=null)
                                {
                                    familyCostInfoClass.cost = Convert.ToDouble(rootObject.records[i].fields.UnitPrice.ToString());
                                    familyCostInfoClass.Remarks = "Highest Unit Price From Airtable"; ;
                                }
                                else
                                {
                                    familyCostInfoClass.cost = 0;
                                    familyCostInfoClass.Remarks = "Unit Price Not provided in Airtable";
                                }
                                if(rootObject.records[i].fields.GenericName!=null)
                                {
                                    familyCostInfoClass.familyGenericName = rootObject.records[i].fields.GenericName.ToString();
                                }
                                else
                                {
                                    familyCostInfoClass.familyGenericName = "-";
                                }
                                //if (rootObject.records[i].fields.Category != null)
                                //{
                                //    familyCostInfoClass.Category = rootObject.records[i].fields.Category.ToString();
                                //}
                                //else
                                //{
                                //    familyCostInfoClass.Category = "-";
                                //}

                                if (rootObject.records[i].fields.Subcategory != null)
                                {
                                    familyCostInfoClass.SubCategory = rootObject.records[i].fields.Subcategory.ToString();
                                }
                                else
                                {
                                    familyCostInfoClass.SubCategory = "-";
                                }

                                if (rootObject.records[i].fields.approvalStatus!=null)
                                {
                                    familyCostInfoClass.approvalStatus = rootObject.records[i].fields.approvalStatus.ToString();
                                }
                                else
                                {
                                    familyCostInfoClass.approvalStatus= "-";
                                }
                              
                                familyCostInfoClassList.Add(familyCostInfoClass);
                            }
                            catch (Exception)
                            {

                                continue;
                            }
                        }
                        // Check if there are more records to retrieve
                        hasMoreRecords = !string.IsNullOrEmpty(rootObject.Offset);
                        offset = rootObject.Offset;
                    }



                    else
                    {
                        Console.WriteLine("Failed to retrieve data from Airtable. Error: " + response.StatusCode);
                        // return null;
                        return familyCostInfoClassList;
                    }



                }
                    return familyCostInfoClassList;  
               
            }
        }

        public static void ListToExcel(List<FamilyInstanceBOQClass> familyInstanceBOQClassList, string FilePath,string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //start excel
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.DefaultRowHeight = 12;
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;
            workSheet.Cells[1, 1].Value = "S.No";
            workSheet.Cells[1, 2].Value = "Type Name";
            workSheet.Cells[1, 3].Value = "Family Name";
            workSheet.Cells[1, 4].Value = "Category";
            workSheet.Cells[1, 5].Value = " Sub Category";
            workSheet.Cells[1, 6].Value = "Generic Name";
            workSheet.Cells[1, 7].Value = "Count";
            workSheet.Cells[1, 8].Value = "Unit Price";
            workSheet.Cells[1, 9].Value = "Total Cost";
            workSheet.Cells[1, 10].Value = " AirTable Approval Status";
            workSheet.Cells[1, 11].Value = "Remarks";
            int recordIndex = 2;
            foreach (var item in familyInstanceBOQClassList)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = item.FamilyType;
                workSheet.Cells[recordIndex, 3].Value = item.FamilyName;
                workSheet.Cells[recordIndex, 4].Value = item.Category;
                workSheet.Cells[recordIndex, 5].Value = item.SubCategory;
                workSheet.Cells[recordIndex, 6].Value = item.familyGenericName;
                workSheet.Cells[recordIndex, 7].Value = item.count;
                workSheet.Cells[recordIndex, 8].Value = item.cost / item.count;
                workSheet.Cells[recordIndex, 9].Value = item.cost;
                workSheet.Cells[recordIndex, 10].Value = item.approvalStatus;
                workSheet.Cells[recordIndex, 11].Value = item.Remarks;
                recordIndex++;
            }
            String iRowCntActual = (workSheet.Dimension.End.Row).ToString();
            String iRowCnt = (workSheet.Dimension.End.Row + 2).ToString();
            string actualPriceEndCellIndex = "I" + iRowCntActual;
            string priceEndCellIndex = "I" + iRowCnt;
            string endCellIndex = "K" + iRowCnt;

            workSheet.Cells[priceEndCellIndex].Formula = string.Format("SUM(F2:{0})", actualPriceEndCellIndex);
            workSheet.Cells[Convert.ToInt32(iRowCnt), 1].Value = "Total";
            using (ExcelRange Rng = workSheet.Cells[Convert.ToInt32(iRowCnt), 1, Convert.ToInt32(iRowCnt), 11])
            {
                //Rng.Merge = true;
                Rng.Style.Font.Bold = true;
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.GreenYellow);

            }
            using (ExcelRange Rng = workSheet.Cells[1, 1,1,11])
            {
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine);
            }

                using (ExcelRange Rng = workSheet.Cells[1, 1, Convert.ToInt32(iRowCnt), 11])
            {
                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            workSheet.Cells[string.Format("A1:{0}", endCellIndex)].AutoFitColumns();
            if (!File.Exists(FilePath +  fileName))
            {
                // string path = "C:\\Users\\Rounik Chatterjee\\Desktop\\New folder (2)\\boq.xlsx";
                Stream stream = File.Create(FilePath + fileName);
                excel.SaveAs(stream);
                stream.Close();
            }
            else
            {
                TaskDialog.Show("Error Message", "Same Name File already exists in the path");
            }

        }

        public static TaskDialogShowClass readExcel(string filePath,string fileName)
        {
            TaskDialogShowClass taskDialogShowClass = new TaskDialogShowClass();
            FileInfo existingFile = new FileInfo(filePath+"\\" + fileName);
           
            using (ExcelPackage package = new ExcelPackage(existingFile ))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet =  package.Workbook.Worksheets["Sheet1"];
                string rowCount = worksheet.Dimension.End.Row.ToString();
                //worksheet.Cells["I27"].Formula = "=COUNT('" + package.Workbook.Worksheets[2].Name + "'!" + package.Workbook.Worksheets[2].Cells["A1:B25"] + ")";

                //calculate all the values of the formulas in the Excel file
                try
                {
                    package.Workbook.Calculate();
                    taskDialogShowClass.totalCost = worksheet.Cells["I" + rowCount].Value.ToString();
                }
                catch (Exception)
                {

                    taskDialogShowClass.totalCost = 0.ToString();
                }




                

            }
            return taskDialogShowClass;
        }

        public static string getCategory(Element element)
        {
            return element.Category.Name;
        }

        public static string getFamilyName(Element element)
        {
            return (element as FamilyInstance).Symbol.FamilyName;
            
        }
    }
}




