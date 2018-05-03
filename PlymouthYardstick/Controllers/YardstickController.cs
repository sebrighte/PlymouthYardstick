using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab
using System.Web.Http.Results;
using System.Web.Http.Description;

namespace PlymouthYardstick.Controllers
{
    public class clsBoat
    {
        public string Index;
        public string ClassName;
        public string Crew;
        public string Rig;
        public string Spinnaker;
        public string PYS;
        public string Change;
        //public string Races;
        //public string Notes;

        public bool SearchForBoat(string strSearch)
        {
            if(this.ClassName.ToLower()==strSearch.ToLower())
                return true;
            else
                return false;
        }

        public int GetAdjustedTime(double RaceTimeSeconds)
        {
            int pys = Int32.Parse(this.PYS);
            return Convert.ToInt32((RaceTimeSeconds / pys) * 1000);
        }
    }

    /// <summary>
    /// Retrieve and allow calculations on data from the RYA PYS ecxel file
    /// </summary>
    public class YardstickController : ApiController
    {
        List<clsBoat> BoatList = new List<clsBoat>();
        YardstickController()
        {
            string path = System.Web.HttpContext.Current.Request.MapPath("~\\Data\\PNLIST.xlsx");

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ernie\Downloads\PNLIST.xlsx");
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"http://www.rya.org.uk/SiteCollectionDocuments/technical/Web%20Documents/PY%20Documentation/PN_List%20-%202018.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int startList = 0, endList = 0;

            for (int i = 1; i <= rowCount; i++)
            {
                try
                {
                    if (xlRange.Cells[i, 1].Value2.ToString() == "Class Name")
                    {
                        startList = i + 1;
                        break;
                    }
                }
                catch { }
            }

            for (int i = 1; i <= rowCount; i++)
            {
                try
                {
                    string a = xlRange.Cells[i, 1].Value2.ToString();

                    //if (xlRange.Cells[i, 1].Value2.ToString() == "EXPERIMENTAL NUMBERS")
                    if (a.Contains("The RYA would like to further thank"))
                    {
                        endList = i - 1;
                        break;
                    }
                }
                catch { }
            }

            int boatIndex = 1;

            for (int i = startList; i <= endList; i++)
            {
                try
                {
                    if (xlRange.Cells[i, 1].Value2.ToString() != "Class Name")
                    {
                        clsBoat Boat = new clsBoat();
                        Boat.Index = boatIndex.ToString();
                        Boat.ClassName = xlRange.Cells[i, 1].Value2.ToString();
                        Boat.Crew = xlRange.Cells[i, 2].Value2.ToString();
                        Boat.Rig = xlRange.Cells[i, 3].Value2.ToString();
                        Boat.Spinnaker = xlRange.Cells[i, 4].Value2.ToString();
                        Boat.PYS = xlRange.Cells[i, 5].Value2.ToString();
                        Boat.Change = xlRange.Cells[i, 6].Value2.ToString();
                        //Boat.Races = xlRange.Cells[i, 7].Value2.ToString();
                        //Boat.Notes = xlRange.Cells[i, 8].Value2.ToString();
                        BoatList.Add(Boat);
                        boatIndex = boatIndex + 1;
                    }
                }
                catch { }
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }


        /// <summary>
        /// Get a list of all boats and PYS details
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Yardstick/GetAllBoats")]
        [ResponseType(typeof(List<clsBoat>))]
        public IHttpActionResult GetAllBoats()
        {
            return Ok(BoatList);
        }

        /// <summary>
        /// Get a list of all boats and PYS details
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Yardstick/GetAllBoatsHTML")]
        public IHttpActionResult GetAllBoatsHTML()
        {
            string HTML = "<table><tr><th>Index</th><th>ClassName</th><th>No.Crew</th><th>Rig</th><th>Spinnaker</th><th>Handicap</th><th>Change in year</th></tr>";
            foreach (clsBoat Boat in BoatList)
            {
                HTML += "<tr><td>" + Boat.Index + "</td>";
                HTML += "<td>" + Boat.ClassName + "</td>";
                HTML += "<td>" + Boat.Crew + "</td>";
                HTML += "<td>" + Boat.Rig + "</td>";
                HTML += "<td>" + Boat.Spinnaker + "</td>";
                HTML += "<td>" + Boat.PYS + "</td>";
                HTML += "<td>" + Boat.Change + "</td></tr>";
            }
            HTML += "</table>";
            return Ok(HTML);
        }


        /// <summary>
        /// Get a list of all PYS Class Names
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Yardstick/GetBoatClassNames")]
        [ResponseType(typeof(List<string>))]
        public IHttpActionResult GetBoatClassNames()
        {
            List<string> iList = new List<string>();
            foreach (clsBoat boat in BoatList)
            {
                iList.Add(boat.ClassName);
            }
            return Ok(iList);
        }

        /// <summary>
        /// Get PYS Class details passing class name
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Yardstick/GetBoatDetails/{ClassName}")]
        [ResponseType(typeof(clsBoat))]
        public IHttpActionResult GetBoatDetails(string ClassName)
        {
            foreach (clsBoat boat in BoatList)
            {
                if (boat.SearchForBoat(ClassName))
                    return Ok(boat);
            }
            return NotFound();
        }

        /// <summary>
        /// Get adjusted race time for class
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Yardstick/GetAdjustedTime")]
        [ResponseType(typeof(int))]
        public IHttpActionResult GetAdjustedTime(string BoatClass, double RaceTimeSeconds)
        {
            foreach (clsBoat boat in BoatList)
            {
                if (boat.SearchForBoat(BoatClass))
                    return Ok(boat.GetAdjustedTime(RaceTimeSeconds));
            }
            return NotFound();
        }
    }
}
