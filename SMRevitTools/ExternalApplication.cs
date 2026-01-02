using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace SMRevitTools
{
    class ExternalApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {

            string path = Assembly.GetExecutingAssembly().Location;

            string tabName = "SM Tools";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { }
            

            //RibbonPanel ProjectTools = application.CreateRibbonPanel("SM Tools", "Project Tools");
            RibbonPanel ProjectSetup = application.CreateRibbonPanel("SM Tools", "🥷 Setup");
            RibbonPanel ModelTools = application.CreateRibbonPanel("SM Tools", "💅 Model Automations");
            RibbonPanel AnnoTools = application.CreateRibbonPanel("SM Tools", "✍️ Annotation Automations");


            #region Document Builder Tool


            PushButtonData buildDocumentButtonData = new PushButtonData("Build Document Data", "🗎 Doc", path, "SMRevitTools.BuildDocument");
            buildDocumentButtonData.LongDescription = "Build Document"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2024.3";
            PushButton buildDocumentButton = ProjectSetup.AddItem(buildDocumentButtonData) as PushButton;
            buildDocumentButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.doc25.png");
            buildDocumentButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.doc50.png");
            #endregion



            #region Enlarged Sheet Tool

            PushButtonData enlargedtButtonData = new PushButtonData("Enlarged Sheet", "EnlargedSheet", path, "SMRevitTools.EnlargedDrawing");
            enlargedtButtonData.LongDescription = "Enlarged Sheets"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2024.3";
            PushButton enlargedSheetButton = ProjectSetup.AddItem(enlargedtButtonData) as PushButton;
            enlargedSheetButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.enlargedDrawing24.png");
            enlargedSheetButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.enlargedDrawing96.png");

            #endregion


            //#region Room Fill Tool

            //PushButtonData roomFillButtonData = new PushButtonData("Room Fill", "Room Fill", path, "SMRevitTools.RoomFill");
            //roomFillButtonData.LongDescription = "Room Fill"
            //            + Environment.NewLine + Environment.NewLine +
            //            "Author: VDC, Space Matrix" + Environment.NewLine +
            //            "Tool Version: 2024.3";
            //PushButton roomFillButton = ProjectTools.AddItem(roomFillButtonData) as PushButton;
            //roomFillButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.roomFill24.png");
            //roomFillButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.roomFill50.png");

            //#endregion

            #region Floor Finish Tool

            PushButtonData floorFinishButtonData = new PushButtonData("Floor Finish", "RoomFloor", path, "SMRevitTools.FloorFinish");
            floorFinishButtonData.LongDescription = "Floor Finish"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025.2";
            PushButton floorFinishButton = ModelTools.AddItem(floorFinishButtonData) as PushButton;
            floorFinishButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.carpet24.png");
            floorFinishButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.carpet64.png");

            #endregion

            #region Ceiling Tool

            PushButtonData ceilingButtonData = new PushButtonData("Ceiling Matrix", "FloatingGrid", path, "SMRevitTools.CeilingMatrix");
            ceilingButtonData.LongDescription = "Ceiling Matrix"
                        + Environment.NewLine + Environment.NewLine +
                        "Creates standard 600x600 grid ceiling with floating gypsum band." +
                        Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton ceilingButton = ModelTools.AddItem(ceilingButtonData) as PushButton;
            ceilingButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.ceilinggyp24.png");
            ceilingButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.ceilinggyp50.png");

            PushButtonData ceilingButtonData2 = new PushButtonData("Ceiling Matrix 2", "GypsumGrid", path, "SMRevitTools.CeilingMatrix2");
            ceilingButtonData2.LongDescription = "Ceiling Matrix 2"
                        + Environment.NewLine + Environment.NewLine +
                        "Creates standard 600x600 grid ceiling with gypsum band both at 2550mm elevation from room level." +
                        Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton ceilingButton2 = ModelTools.AddItem(ceilingButtonData2) as PushButton;
            ceilingButton2.LargeImage = FetchPngIcon("SMRevitTools.Logo.ceiling2_24.png");
            ceilingButton2.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.ceiling2_50.png");
            #endregion




            #region Door Tag Tool

            PushButtonData doorTagButtonData = new PushButtonData("DoorTags", "KnockKnock", path, "SMRevitTools.DoorTagTool");
            doorTagButtonData.LongDescription = "KnockKnock - Tag Doors"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton doorTagButton = AnnoTools.AddItem(doorTagButtonData) as PushButton;
            doorTagButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.door24.png");
            doorTagButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.door100.png");

            #endregion



            //#region Dimension Tool

            //PushButtonData partitionDimButtonData = new PushButtonData("PartDim", "PartDim", path, "SMRevitTools.PartitionDim");
            //partitionDimButtonData.LongDescription = "Dimension the Partitions"
            //            + Environment.NewLine + Environment.NewLine +
            //            "Author: VDC, Space Matrix" + Environment.NewLine +
            //            "Tool Version: 2025(beta)";
            //PushButton partitionDimButton = ProjectTools.AddItem(partitionDimButtonData) as PushButton;
            //partitionDimButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.partDim24.png");
            //partitionDimButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.partDim96.png");

            //#endregion


            #region Room Num Tag Tools

            PushButtonData roomNumberingButtonData = new PushButtonData("RoomNum", "🔟 RoomNum", path, "SMRevitTools.RoomNum");
            roomNumberingButtonData.LongDescription = "Auto Room Numbering"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton roomNumberingmButton = AnnoTools.AddItem(roomNumberingButtonData) as PushButton;
            roomNumberingmButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.roomnum24.png");
            roomNumberingmButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.roomnum64.png");


            PushButtonData roomTagButtonData = new PushButtonData("RoomTag", "🏷️ RoomTag", path, "SMRevitTools.IntRoomTag");
            roomTagButtonData.LongDescription = "Auto Room Tagging"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton roomTagmButton = AnnoTools.AddItem(roomTagButtonData) as PushButton;
            roomTagmButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.roomTag24.png");
            roomTagmButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.roomTag64.png");


            PushButtonData copyRoomTagButtonData = new PushButtonData("CopyRoomTag", "☕ CopyTag", path, "SMRevitTools.CopyRoomTag");
            copyRoomTagButtonData.LongDescription = "Copy Room Tags to other views"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, Space Matrix" + Environment.NewLine +
                        "Tool Version: 2025(beta)";
            PushButton copyRoomTagmButton = AnnoTools.AddItem(copyRoomTagButtonData) as PushButton;
            copyRoomTagmButton.LargeImage = FetchPngIcon("SMRevitTools.Logo.copy24.png");
            copyRoomTagmButton.ToolTipImage = FetchPngIcon("SMRevitTools.Logo.copy64.png");



            #endregion



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
