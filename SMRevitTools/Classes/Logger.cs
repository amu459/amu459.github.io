using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using System.IO;

namespace HelperClassLibrary
{
    public class logger
    {
        public static void CreateDump(string toolName, string errorMsg, Document doc, UIApplication uiApp, int timeDelta)
        {
            string gDrivePath = "G:\\Shared drives\\DesignTechnology\\Automations\\Logs Dump\\";
            string localPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                "\\Autodesk\\Revit\\Addins\\";
            //int timeDelta = 0;
            string projectName = doc.Title;
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string userName = uiApp.Application.Username;

            try
            {
                using (StreamWriter writetext = new StreamWriter(gDrivePath + timeStamp + "_" + userName + ".txt"))
                {
                    writetext.WriteLine(timeStamp + ","
                        + userName + ","
                        + toolName + ","
                        + projectName + ","
                        + timeDelta + ","
                        + "Result = "
                        + errorMsg);
                }
            }
            catch
            {
                using (StreamWriter writetext = new StreamWriter(localPath + timeStamp + "_" + userName + ".txt"))
                {
                    writetext.WriteLine(timeStamp + ","
                        + userName + ","
                        + toolName + ","
                        + projectName + ","
                        + timeDelta + ","
                        + "Result = "
                        + errorMsg);
                }
            }

        }

    }
}
