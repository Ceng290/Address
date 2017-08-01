using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OptaAddress
{
    public partial class Default : System.Web.UI.Page
    {
        //**************************************************************************************
        //* Global variables
        //**************************************************************************************
        string savePath;
        string fileName;
        string pathToCheck;
        public static string regexMatch= "";

        //**************************************************************************************
        //* Page_Load Event
        //**************************************************************************************
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //create new session variable 
                Session["CheckRefresh"] = Server.UrlDecode(System.DateTime.Now.ToString());

                //Update the gridview
                updateGridView();

            }else{
                //Do something

            }
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        protected void updateGridView()
        {
            string[] filePaths = Directory.GetFiles(Server.MapPath("~/processedFile/"));
            List<ListItem> files = new List<ListItem>();
            foreach (string filePath in filePaths)
            {
                files.Add(new ListItem(Path.GetFileName(filePath), filePath));
            }
            GridView1.DataSource = files;
            GridView1.DataBind();
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        protected void UploadButton_Click(object sender, EventArgs e)
        {
            //check to see if the page has been refreshed or if a control has been clicked
            if (Session["CheckRefresh"].ToString() == ViewState["CheckRefresh"].ToString())
            {
                // Before attempting to save the file, verify
                // that the FileUpload control contains a file.
                if (FileUpload1.HasFile)
                {
                    // Call a helper method routine to save the file.
                    SaveFile(FileUpload1.PostedFile);

                }else { 
                    // Notify the user that a file was not uploaded.
                    UploadStatusLabel.Text = "You did not specify a file to upload.";
                }

                //reset the session variable
                Session["CheckRefresh"] = Server.UrlDecode(System.DateTime.Now.ToString());
                
            }
            //update the gridview so that it shows all the files in the "~/receivedFiles" folder
            updateGridView();
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        void SaveFile(HttpPostedFile file)
        {
            // Specify the path to save the uploaded file to.
            savePath = Server.MapPath("~/receivedFile/");

            // Get the name of the file to upload.
            fileName = FileUpload1.FileName;

            // Create the path and file name to check for duplicates.
            pathToCheck = savePath + fileName;

            // Create a temporary file name to use for checking duplicates.
            string tempfileName1 = "";
            string tempfileName2 = "";

            // Check to see if a file already exists with the
            // same name as the file to upload.        
            if (System.IO.File.Exists(pathToCheck))
            {
                int counter = 2;
                while (System.IO.File.Exists(pathToCheck))
                {
                    // if a file with this name already exists, prefix the filename with a number.
                    tempfileName1 = counter.ToString() + "-" + fileName;
                    pathToCheck = savePath + tempfileName1;
                    counter++;
                }
                tempfileName2 = tempfileName1;
                tempfileName2 = tempfileName2.Replace("xlxs", "xls");
                fileName = tempfileName1;

                // Notify the user that the file name was changed.
                UploadStatusLabel.Text = "A file with the same name already exists.  <br />Your file was saved as " + tempfileName2;

            }else{
                // Notify the user that the file was saved successfully.
                UploadStatusLabel.Text = "Your Excel file was uploaded, and updated successfully.";
            }

            //Append the name of the file to upload to the path.
            savePath += fileName;

            // Call the SaveAs method to save the uploaded
            // file to the specified directory.
            FileUpload1.SaveAs(savePath);

            //Here is where we parse the file 
            updateFile(ref pathToCheck);


            //Update the gridview
            updateGridView();
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        void updateFile(ref String filePath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            string parsedString;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    //Here is where we parse the string in the cell with "|"
                    parsedString = parseCellValue(ref str);
                    //update the cell with new parsedString
                    xlWorkSheet.Cells[rCnt, cCnt] = parsedString;                                     
                }
            }

            //update the folder that the file is to be saved to
            filePath = filePath.Replace("receivedFile", "processedFile");
            //Update the file extension from xlsx to xls so that the file can be opened without errors
            filePath = filePath.Replace("xlsx", "xls");
            //save the file
            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //close the workbook
            xlWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);

            //set the book and worksheet to null
            xlWorkSheet = null;
            xlWorkBook = null;

            xlApp.Quit();

            //Garbage collection
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string parseCellValue(ref string str) {

            //Remove white space from start of string
            str = str.TrimStart();
            //Remove white space from end of string
            str = str.TrimEnd();
            //Remove all commas
            str = str.Replace(",", " ");
            //Assign value to regexMatch
            regexMatch = str;

            //Create a dictionary to parse the regexMatch into
            Dictionary<int, string> dictionary = new Dictionary<int, string>();
            dictionary[1] = getStreetNumber(ref regexMatch);
            dictionary[15] = getPostalCode(ref regexMatch);
            dictionary[13] = getProvince(ref regexMatch);
            dictionary[9] = getStreetUnitApt(ref regexMatch);
            dictionary[7] = getStreetDirection(ref regexMatch);
            dictionary[5] = getStreetType(ref regexMatch);
            dictionary[3] = getStreetName(ref regexMatch);
            dictionary[11] = getStreetMunicipality(ref regexMatch);
            dictionary[2] = "|";
            dictionary[4] = "|";
            dictionary[6] = "|";
            dictionary[8] = "|";
            dictionary[10] = "|";
            dictionary[12] = "|";
            dictionary[14] = "|";

            // Acquire keys and sort them.
            var list = dictionary.Keys.ToList();
            list.Sort();

            string toString = "";

            // Loop through keys and build the string to append to str
            foreach (var key in list)
            {
                toString = toString + dictionary[key].Trim();
            }

            //Add | to end of string
            str = str + "|";

            //Conjoin toString and Str to get final string
            str = str + toString;

            return str;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetNumber(ref string str)
        {
            string number = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            //REGEX Patterns
            string streetNumberPattern = "^[0-9]*\\s";

            Match match = Regex.Match(str, streetNumberPattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                number = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(number, "  ");
            }

            return number;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetName(ref string str)
        {
            string name = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            //REGEX Pattern
            string namePatern = "^(.*)\\s\\s";

            Match match1 = Regex.Match(str, namePatern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match1.Success)
            {
                // Finally, we get the Group value and display it.
                name = match1.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(name, "");
            }

            return name;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetType(ref string str)
        {
            string type = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            //REGEX Patterns
            string a1 = "\\s(ABBEY)\\s|\\s(ACRES)\\s|\\s(ALLÉE)\\s|\\s(ALLEY)\\s|\\s(AUT)\\s|\\s(AVE?.)\\s|\\s(AV)\\s|\\s(BAY)\\s|\\s(BEACH)\\s|\\s(BEND)\\s|\\s(BLVD)\\s|\\s(BOUL)\\s";
            string a2 = "|\\s(BYPASS)\\s|\\s(BYWAY)\\s|\\s(CAMPUS)\\s|\\s(CAPE)\\s|\\s(CAR)\\s|\\s(CARREF)\\s|\\s(CTR?.)\\s|\\s(CERCLE)\\s|\\s(CHASE)\\s|\\s(CH?.)\\s|\\s(CIR)\\s";
            string a3 = "|\\s(CIRCT)\\s|\\s(CLOSE)\\s|\\s(COMMON)\\s|\\s(CONC)\\s|\\s(CRNRS)\\s|\\s(CÔTE)\\s|\\s(COUR)\\s|\\s(COURS)\\s|\\s(CRT?.)\\s|\\s(COVE)\\s|\\s(CRES?.)\\s";
            string a4 = "|\\s(CROIS)\\s|\\s(CROSS)\\s|\\s(CDS?.)\\s|\\s(C?.)\\s|\\s(DALE)\\s|\\s(DELL)\\s|\\s(DIVERS)\\s|\\s(DOWNS)\\s|\\s(DR?.)\\s|\\s(ÉCH)\\s|\\s(END)\\s|\\s(ESPL)\\s|\\s(ESTATE)\\s";
            string a5 = "|\\s(EXPY)\\s|\\s(EXTEN)\\s|\\s(FARM)\\s|\\s(FIELD)\\s|\\s(FOREST)\\s|\\s(FWY?.)\\s|\\s(FRONT)\\s|\\s(GDNS)\\s|\\s(GATE)\\s|\\s(GLADE)\\s|\\s(GLEN?.)\\s";
            string a6 = "|\\s(GREEN)\\s|\\s(GRNDS)\\s|\\s(GROVE)\\s|\\s(HARBR)\\s|\\s(HEATH)\\s|\\s(HTS?.)\\s|\\s(HGHLDS)\\s|\\s(HWY?.)\\s|\\s(HILL)\\s|\\s(HOLLOW)\\s|\\s(ÎLE?.)\\s";
            string a7 = "|\\s(IMP?.)\\s|\\s(INLET)\\s|\\s(ISLAND)\\s|\\s(KEY?.)\\s|\\s(KNOLL)\\s|\\s(LANDNG)\\s|\\s(LANE)\\s|\\s(LMTS)\\s|\\s(LINE)\\s|\\s(LINK)\\s|\\s(LKOUT)\\s";
            string a8 = "|\\s(LOOP)\\s|\\s(MALL)\\s|\\s(MANOR)\\s|\\s(MAZE)\\s|\\s(MEADOW)\\s|\\s(MEWS)\\s|\\s(MONTÉE)\\s|\\s(MOOR)\\s|\\s(MOUNT)\\s|\\s(MTNORCH)\\s|\\s(PARADE)\\s";
            string a9 = "|\\s(PARC?.)\\s|\\s(PK?.)\\s|\\s(PKY?.)\\s|\\s(PASS)\\s|\\s(PATH)\\s|\\s(PTWAY)\\s|\\s(PINES)\\s|\\s(PL?.)\\s|\\s(PLACE)\\s|\\s(PLAT)\\s|\\s(PLAZA)\\s|\\s(PT?.)\\s|\\s(POINTE)\\s";
            string a10 = "|\\s(PORT)\\s|\\s(PVT?.)\\s|\\s(PROM)\\s|\\s(QUAI)\\s|\\s(QUAY)\\s|\\s(RAMP)\\s|\\s(RANG)\\s|\\s(RG?.)\\s|\\s(RIDGE)\\s|\\s(RISE)\\s|\\s(RD?.)\\s|\\s(RDPT)\\s|\\s(RTE?.)\\s";
            string a11 = "|\\s(ROW)\\s|\\s(RUE?.)\\s|\\s(RLE?.)\\s|\\s(RUN)\\s|\\s(SENT)\\s|\\s(SQ?.)\\s|\\s(STREET)\\s|\\s(ST?.)\\s|\\s(SUBDIV)\\s|\\s(TERR?.)\\s|\\s(TSSE?.)\\s|\\s(THICK)\\s|\\s(TOWERS)\\s|\\s(TLINE)\\s";
            string a12 = "|\\s(TRAIL)\\s|\\s(TRNABT)\\s|\\s(VALE)\\s|\\s(VIA?.)\\s|\\s(VIEW)\\s|\\s(VILLGE)\\s|\\s(VILLAS)\\s|\\s(VISTA)\\s|\\s(VOIE?.)\\s|\\s(WALK)\\s|\\s(WAYWHARF)\\s";
            string a13 = "|\\s(WOOD)\\s|\\s(WYND?.)\\s";

            string streetTypesPattern = a1 + a2 + a3 + a4 + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13;

            Match match = Regex.Match(str, streetTypesPattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                type = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(type, "  ");

            }
            
            return type;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetDirection(ref string str)
        {
            string direction = " ";
            str = str.TrimStart();
            str = str.TrimEnd();
            //REGEX Pattern
            string a = "\\s(E)\\s|\\s(N)\\s|\\s(NE)\\s|\\s(NW)\\s|\\s(S)\\s|\\s(SE)\\s";
            string b = "|\\s(SW)\\s |\\s(W)\\s |\\s(E)\\s |\\s(N)\\s |\\s(N)\\s |\\s(E)\\s";
            string c = "|\\s(NO)\\s |\\s(S)\\s |\\s(SE)\\s |\\s(SO)\\s |\\s(O)\\s";

            string streetDirectionsPattern = a + b + c;


            Match match = Regex.Match(str, streetDirectionsPattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                direction = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(direction, "  ");
            }

            return direction;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetUnitApt(ref string str)
        {
            string unitApt = " ";
            str = str.TrimStart();
            str = str.TrimEnd();
            //REGEX Pattern
            string unitDesignatorsPattern = "\\s(APT)\\s|\\s(SUITE)\\s|\\s(UNIT)\\s|\\s(APP)\\s|\\s(BUREAU)\\s|\\s(UNITÉ)\\s";

            Match match = Regex.Match(str, unitDesignatorsPattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                unitApt = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(unitApt, "  ");
            }

            return unitApt;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getStreetMunicipality(ref string str)
        {
            string municipality = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            municipality = str;
            regexMatch = regexMatch.Replace(municipality, "");

            return municipality;
        }
        
        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getProvince(ref string str)
        {
            string province = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            //REGEX Patterns
            string a = "\\s(AB)$|\\s(BC)$|\\s(MB)$|\\s(NB)$|\\s(NL)$|\\s(NT)$|\\s(NS)$|\\s(NU)$|\\s(ON)$|\\s(PE)$|\\s(QC)$|\\s(SK)$|\\s(YT)$";
            string b = "|\\s(Alberta)$|\\s(British\\sColumbia)$|\\s(Manitoba)$|\\s(New\\sBrunswick)$|\\s(Newfoundland\\sand\\sLabrador)$|\\s(Northwest\\sTerritories)$|\\s(Nova\\sScotia)$|\\s(Nunavut)$|\\s(Ontario)$|\\s(Prince\\sEdward\\sIsland)$|\\s(Québec)$|\\s(Saskatchewan)$|\\s(Yukon)$";
            string c = "|\\s(Colombie-Britannique)$|\\s(Manitoba)$|\\s(Nouveau-Brunswick)$|\\s(Terre-Neuve-et-Labrador)$|\\s(Territoires\\sdu\\sNord-Ouest)$|\\s(Nouvelle-Écosse)$|\\s(Nunavut)$|\\s(Île-du-Prince-Édouard)$";
            string d = "|\\s(NL)$|\\s(PEI)$|\\s(QUE)$|\\s(ONT)$|\\s(MAN)$|\\s(SASK)$|\\s(ALTA)$|\\s(NWT)$|\\s(NVT)$|\\s(TNO)$|\\s(CB)$|\\s(ALB)$|\\s(NÉ)$|\\s(QUE)$|\\s(Quebec)$";

            string provinceTerritoriesPattern = a + b + c + d;

            Match match = Regex.Match(str, provinceTerritoriesPattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                province = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(province, "  ");
            }

            return province;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        static string getPostalCode(ref string str)
        {
            string postalCode = " ";
            str = str.TrimStart();
            str = str.TrimEnd();

            //REGEX Pattern
            string postalCodePattern = "[ABCEGHJKLMNPRSTVXY][0-9][ABCEGHJKLMNPRSTVWXYZ] ?[0-9][ABCEGHJKLMNPRSTVWXYZ][0-9]";

            Match match = Regex.Match(str, postalCodePattern, RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                postalCode = match.Value;

                //update the regexMatch String by removing the matched string
                regexMatch = regexMatch.Replace(postalCode, " ");
            }

            return postalCode;
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        protected void DownloadFile(object sender, EventArgs e)
        {
            string filePath = (sender as LinkButton).CommandArgument;
            Response.ContentType = ContentType;
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(filePath));
            Response.WriteFile(filePath);
            Response.End();
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        protected void DeleteFile(object sender, EventArgs e)
        {
            try {
                string filePath = (sender as LinkButton).CommandArgument;
                File.Delete(filePath);
                Response.Redirect(Request.Url.AbsoluteUri);

            } catch {

            }           
        }

        //**************************************************************************************
        //*
        //**************************************************************************************
        protected void Page_PreRender(object sender, EventArgs e)
        {
            ViewState["CheckRefresh"] = Session["CheckRefresh"];
        }
    }
}