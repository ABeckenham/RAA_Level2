#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Level = Autodesk.Revit.DB.Level;
//using Excel = Microsoft.Office.Interop.Excel;



#endregion

namespace RAA_Level2
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // put any code needed for the form here



            //collect for the titleblock
            FilteredElementCollector titleblockcollector = new FilteredElementCollector(doc);
            titleblockcollector.OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType();

            //collect viewfamilytype of each view to find floorplans and ceilings
            FilteredElementCollector viewtypecollector = new FilteredElementCollector(doc);
            viewtypecollector.OfClass(typeof(ViewFamilyType));

            //collect the ceiling and floorplan view types
            Dictionary<ViewFamily, ViewFamilyType> viewFamilyTypesDic = new Dictionary<ViewFamily, ViewFamilyType>();


            foreach (ViewFamilyType curType in viewtypecollector)
            {
                if (!viewFamilyTypesDic.ContainsKey(curType.ViewFamily))
                {
                    //viewFamilyTypesDic.Add(curType.ViewFamily, curType);
                    //above is the method i learnt but below is the short hand version which is dictionary[key]=value
                    viewFamilyTypesDic[curType.ViewFamily] = curType;
                }

                if (viewFamilyTypesDic.Count == Enum.GetValues(typeof(ViewFamily)).Length)
                {
                    break; //exit once all viewTypes have been found
                }
            }

            //get ceiling and floor plans viewfamilytypes
            ViewFamilyType viewCeilingPan = null;
            ViewFamilyType viewFloorPlan = null;

            if (viewFamilyTypesDic.TryGetValue(ViewFamily.FloorPlan, out ViewFamilyType floorPlan))
            {
                viewFloorPlan = floorPlan;
            }

            if (viewFamilyTypesDic.TryGetValue(ViewFamily.CeilingPlan, out ViewFamilyType ceilingPlan))
            {
                viewCeilingPan = ceilingPlan;
            }

            if (viewFloorPlan == null || viewCeilingPan == null)
            {
                TaskDialog.Show("Error", "Unable to find both floor plan and ceiling plan view types.");
            }

            //////////////////////FORM//////////////////////

            //2. open form
            MyForm currentForm = new MyForm()
            {
                Width = 500,
                Height = 400,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            currentForm.ShowDialog();

            //4. get form data and do something
            if (currentForm.DialogResult == false)
            {
                return Result.Cancelled;
            }

            //get path to excel file from form
            string filepath = currentForm.getTextboxValue();
            bool makefloorplans = currentForm.getChbFloorPlans();
            bool makeCeilingPlan = currentForm.getChbCeilingPlans();
            string currentunit = currentForm.getUnitGroup();

            #region InteropExcel
            ////////////////////EXCEL///////////////////////

            ////open excel file
            //Excel.Application excelAp = new Excel.Application();
            //Excel.Workbook excelworkb = excelAp.Workbooks.Open(filepath);
            //Excel.Worksheet worksheet = excelworkb.Worksheets[1];
            //Excel.Range range = worksheet.UsedRange;

            ////get rows and columns of data 
            //int rows = range.Rows.Count;
            //int cols = range.Columns.Count;

            ////write them to list
            //List<List<string>> exceldata = new List<List<string>>();
            //List<string> Levellist = new List<string>();
            //List<string> elevationMlist = new List<string>();
            //List<string> elevationFlist = new List<string>();

            ////read data into list
            //for (int i = 1; i <= rows; i++)
            //{
            //    List<string> rowdata = new List<string>();
            //    for (int j = 1; j <= cols; j++)
            //    {
            //        string colData = Convert.ToString(worksheet.Cells[i, j].Value);
            //        rowdata.Add(colData);

            //        if (j == 1)
            //        {
            //            Levellist.Add(colData);
            //        }
            //    }
            //    exceldata.Add(rowdata); // level name, elevation in meters, elevation in foot

            //}
            //Levellist.RemoveAt(0);


            ////close the excel!
            //excelworkb.Save();
            //excelAp.Quit();

            ////////////
            #endregion InteropExcel


            List<List<string>> excelData = new List<List<string>>();
            List<string> levelList = new List<string>();

            //open excel using OPENXML instead Interop keeps crashing
            ExcelDataManager excelManager = new ExcelDataManager(filepath);

            string sheetName = "Sheet1";                   
            

            IReadOnlyList<IReadOnlyList<string>> worksheetData = excelManager.GetWorksheet(sheetName);
            int listcount = worksheetData.Count;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Created plan on levels");                

                //create levels         
                for(int i =0; i<listcount;i++)
                {
                    if (i == 0)
                    {
                        continue;
                    }
                    else
                    {

                        string levelName = excelManager.GetCellValue(sheetName, i, 0);
                        double currentheight = 0;
                        if (currentunit == "Imperial")
                        {
                            string currentheightstring = excelManager.GetCellValue(sheetName, i, 1);
                            currentheight = Convert.ToDouble(currentheightstring);
                        }
                        else
                        {
                            string currentheightstring = excelManager.GetCellValue(sheetName, i, 2);
                            currentheight = Convert.ToDouble(currentheightstring);
                        }
                        Level newLevel = Level.Create(doc, currentheight);
                        newLevel.Name = levelName;

                        if (makefloorplans == true)
                        {
                            //create floor plans
                            ViewPlan curFloorPlan = ViewPlan.Create(doc, viewFloorPlan.Id, newLevel.Id);
                            curFloorPlan.Name = levelName + "_Floor Plan";

                        }
                        if (makeCeilingPlan == true)
                        {

                            //create ceiling plans
                            ViewPlan curCeilingplan = ViewPlan.Create(doc, viewCeilingPan.Id, newLevel.Id);
                            curCeilingplan.Name = newLevel.Name.ToString() + "_Ceiling Plan";
                        }
                    }
                }

                t.Commit();

            }

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

        public double convertFeetToMetres(double feet)
        {
            return feet / 3.2808399;
        }


    }
}