using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace DeskAutomation
{
    class ExternalApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string path = assembly.Location;

            #region Checking Tab and Creating DeskAutomation Panel
            string tabName = "India Tools";
            string panelName = "一═デ︻  Layout Machine  ︻デ═一";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { }
            List<RibbonPanel> panelList = application.GetRibbonPanels(tabName);
            RibbonPanel panelDA = null;
            foreach (RibbonPanel rp in panelList)
            {
                if (rp.Name == panelName)
                {
                    panelDA = rp;
                }
            }
            if (panelDA == null)
            {
                panelDA = application.CreateRibbonPanel(tabName, panelName);
            }
            #endregion
            panelDA.AddSeparator();

            PushButtonData helloRevit = new PushButtonData("Hello Revit Button",
                "r/AutoDesk", path, "DeskAutomation.HelloRevit");
            helloRevit.LargeImage = FetchPngIcon("DeskAutomation.icons.surfRoboZ.png");

            PushButton helloRevitButton = panelDA.AddItem(helloRevit) as PushButton;
            helloRevitButton.ToolTipImage = FetchPngIcon("DeskAutomation.icons.surfRobo256.png");
            helloRevitButton.LongDescription = "Desk layout according to door location and orientation"
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.2 or higher" + Environment.NewLine +
                "2. Load Standard Desk family in project " + Environment.NewLine +
                "3. Physical Desk Family Name is '1_Person-Office-Desk'"
                +
                Environment.NewLine + Environment.NewLine + "Steps: " + Environment.NewLine +
                "1. Select room(s) with door on one of the room walls" + Environment.NewLine +
                "2. Click on the tool" + Environment.NewLine +
                "3. Make sure to Manually delete/add/move desks where ever required."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.4";


            PushButtonData VDeskLayout = new PushButtonData("VDesk Layout", 
                "V Layout", path, "DeskAutomation.VDeskLayout");
            VDeskLayout.Image = FetchPngIcon("DeskAutomation.icons.vLayout50.png");

            //PushButton VDeskButton = panelDA.AddItem(VDeskLayout) as PushButton;
            VDeskLayout.ToolTipImage = FetchPngIcon("DeskAutomation.icons.verHigh.png");
            VDeskLayout.LongDescription = "Vertical Desk layout -"
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.2 or higher" + Environment.NewLine +
                "2. Load Standard Desk family in project " + Environment.NewLine +
                "3. Physical Desk Family Name is '1_Person-Office-Desk'"
                +
                Environment.NewLine + Environment.NewLine + "Steps: " + Environment.NewLine +
                "1. Select room(s) with or without door; idgaf" + Environment.NewLine +
                "2. Click on the tool" + Environment.NewLine +
                "3. Make sure to Manually delete/add/move desks where ever required."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.4";


            PushButtonData HDeskLayout = new PushButtonData("HDesk Layout",
            "H Layout", path, "DeskAutomation.HDeskLayout");
            HDeskLayout.Image = FetchPngIcon("DeskAutomation.icons.hLayout50.png");

            //PushButton HDeskButton = panelDA.AddItem(HDeskLayout) as PushButton;
            HDeskLayout.ToolTipImage = FetchPngIcon("DeskAutomation.icons.horHigh.png");
            HDeskLayout.LongDescription = "Horizontal Desk layout -"
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.2 or higher" + Environment.NewLine +
                "2. Load Standard Desk family in project " + Environment.NewLine +
                "3. Physical Desk Family Name is '1_Person-Office-Desk'"
                +
                Environment.NewLine + Environment.NewLine + "Steps: " + Environment.NewLine +
                "1. Select room(s) with or without door; idgaf" + Environment.NewLine +
                "2. Click on the tool" + Environment.NewLine +
                "3. Make sure to Manually delete/add/move desks where ever required."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.4";


            PushButtonData AngleDeskLayout = new PushButtonData("Angled Desk Layout",
            "Twist", path, "DeskAutomation.AngleDeskLayout");
            AngleDeskLayout.Image = FetchPngIcon("DeskAutomation.icons.angHigh50.png");

            //PushButton HDeskButton = panelDA.AddItem(HDeskLayout) as PushButton;
            AngleDeskLayout.ToolTipImage = FetchPngIcon("DeskAutomation.icons.angHigh.png");
            AngleDeskLayout.LongDescription = "Desk layout for Rotated Rooms"
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2023.2 or higher" + Environment.NewLine +
                "2. Load Standard Desk family in project " + Environment.NewLine +
                "3. Physical Desk Family Name is '1_Person-Office-Desk'"
                +
                Environment.NewLine + Environment.NewLine + "Steps: " + Environment.NewLine +
                "1. Select single room with or without door; idgaf" + Environment.NewLine +
                "2. Click on the tool" + Environment.NewLine +
                "3. Input the Room angle in degrees" + Environment.NewLine +
                "4. Make sure to Manually delete/add/move desks where ever required."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2023.3";

            panelDA.AddSeparator();

            List<RibbonItem> itemLayouts= new List<RibbonItem>();
            itemLayouts.AddRange(panelDA.AddStackedItems(VDeskLayout, HDeskLayout, AngleDeskLayout));

            panelDA.AddSeparator();

            return Result.Succeeded;
        }

        private System.Windows.Media.ImageSource FetchPngIcon(string embeddedPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
