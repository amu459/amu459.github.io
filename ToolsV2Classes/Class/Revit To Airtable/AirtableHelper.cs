using Autodesk.Revit.DB;
//using HelperClassLibrary.Airtable;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ToolsV2Classes.Class.Revit_To_Airtable
{


    internal class AirtableHelper
    {


        public List<List<string>> Category { get; set; }
        public async void GetRecords()
        {
            string url = "https://api.airtable.com/v0/app3ugEupPfPXDgno/Sample%20BOQ";
            //string bearerToken = "keysVRceC6Li86DSm";
            string bearerToken = "patyZP7S1BUjT5grK.3799b680ad6bbfdf708051db00272695ed9cfdc58733c3684b418be560565368";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + bearerToken);

                HttpResponseMessage response = await client.GetAsync(url);
                string responseBody = await response.Content.ReadAsStringAsync();

                //string header = response.Content.ReadAsStringAsync();

                Console.WriteLine(responseBody);
            }

        }
    }
}
