using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace SMRevitTools
{
    internal class SMSheetClass
    {
        public ViewSheet OriginalSheet { get; set; }

        public string SheetName { get; set; }

        public string SheetNum { get; set; }

        public string SheetCategory { get; set; }

        public string SheetSubCategory { get; set; }

        public string SheetSeries { get; set; }

        public List<View> ViewsInSheet { get; set; }

        public List<View> LegendsInSheet { get; set; }

        public List<ViewSchedule> SchedulesInSheet {  get; set; }

        public List<Viewport> ViewportsInSheet { get; set; }


        public void GetSheetData(ViewSheet sheet, Document doc, Level level) 
        {
            OriginalSheet = sheet;
            string geometricLevel = level.LookupParameter("SM-Geometric Level").AsString();

            SheetName = level.Name + " " + sheet.Name.Remove(0, 8);
            //SheetNum = sheet.SheetNumber.Substring(0, 7) + geometricLevel;
            SheetNum = sheet.SheetNumber.Replace("F00", "F"+ geometricLevel);
            SheetCategory = sheet.LookupParameter("SM-SheetCategory").AsString();
            SheetSubCategory = sheet.LookupParameter("SM-SheetSubCategory").AsString();

            SheetSeries = sheet.LookupParameter("SM-SheetSeries").AsString();
            if(SheetSeries.ToLower().Contains("x floor"))
            {
                SheetSeries = geometricLevel + "-" + level.Name;
            }
            List<View> tempViewList = new List<View>();
            List<View> tempLegendList = new List<View>();
            var combinedViewList = sheet.GetAllPlacedViews().ToList();


            if (combinedViewList.Count() > 0)
            {
                foreach (var viewId in combinedViewList)
                {
                    View v = doc.GetElement(viewId) as View;
                    if (v != null)
                    {
                        if (v.ViewType == ViewType.Legend)
                        {
                            tempLegendList.Add(v);
                        }
                        else
                        {
                            tempViewList.Add(v);
                        }
                    }
                }
            }
            LegendsInSheet = tempLegendList;
            ViewsInSheet = tempViewList;

            List<Viewport> tempViewportList = new List<Viewport>();
            var viewportIds = sheet.GetAllViewports().ToList();
            if(viewportIds.Count > 0)
            {
                foreach(var viewportId in  viewportIds)
                {
                    Viewport tempViewport = doc.GetElement(viewportId) as Viewport;
                    if(tempViewport != null)
                    {
                        tempViewportList.Add(tempViewport);
                    }
                }
            }
            ViewportsInSheet = tempViewportList;
        }
    }
}
