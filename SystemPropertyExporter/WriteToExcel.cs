using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Internal.ApiImplementation;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Plugins;
using System.Collections.ObjectModel;
using SystemPropertyExporter;

namespace SystemPropertyExporter
{
    class WriteToExcel
    {
        //REQUIRED TO BE GLOBAL PARAMETER SO A RUNNING TOTAL CAN BE KEPT
        //PREVENTS HAVING TO CYCLE THROUGH ENTIRE 0 BASED ItemIdx List in RangeProp METHOD
        public static int IdxCounter { get; set; }

        //TRANSFERS ExportItems TO EXCEL FILE FOR USER PURPOSES
        public static void ExcelReport()
        {
            try
            {
                //CREATE NEW INTANCE OF EXCEL APPLICATION TO CREATE NEW WORKBOOK
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook;

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!");
                }

                //CREATE NEW WORKBOOK AND WORKSHEETS
                xlWorkbook = xlApp.Workbooks.Add(Missing.Value);
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.Add(
                    Type.Missing,Type.Missing, ExportProperties.UserItems.Count+1, Type.Missing);
                
                //CONSTANTS ASSIGNMENT FOR VALUES TO BE ASSIGNED TO CELLS
                int rowNum = 2;
                int colNum = 7 ;
                int modelIdx = 0;
                IdxCounter = 0;
                bool match = false;

                //-----------------------------------------------------------------------------------------------------
                //BEGIN DATA TRANSFER/CELL RECORDING FROM ExportItems List//
                //ITERATES OVER EACH ITEM
                foreach (Export item in ExportProperties.ExportItems)
                {
                    //1. ITERATES OVER CURRENT SHEET IN WORKBOOK, TO CHECK IF SHEET ALREADY
                    //EXISTS FOR BUILDING SYSTEM/DISCIPLINE
                    foreach (Excel.Worksheet sheet in xlWorkbook.Worksheets)
                    {
                        //IF SHEET ALREADY EXISTS FOR DISCIPLINE,
                        //STARTS RECORDING IN NEXT BLANK ROW OF CELLS
                        if (sheet.Name == item.ExpDiscipline)
                        {
                            match = true;

                            xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[sheet.Name];
                            xlWorksheet.Select();
                            xlWorksheet.Activate();

                            //IF DISCIPLINE SHEET NAME MATCHES, SELECTS NEXT BLANK ROW TO START STORING VALUES
                            Excel.Range last = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                            rowNum = last.Row + 1;
                            break;
                        }
                    }

                    //IF NO WORKSHEET EXISTS FOR CURRENT DISCIPLINE,
                    //CREATES NEW WORKSHEET, ACTIVATES FOR STORING, ASSIGNS SHEET NAME, AND ASSIGNS COLUMN HEADERS 
                    if (match == false)
                    {
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[modelIdx + 1];
                        xlWorksheet.Select();
                        xlWorksheet.Activate();

                        xlWorksheet.Name = item.ExpDiscipline;
                        xlWorksheet.Cells[1, 1] = "DISCIPLINE";
                        xlWorksheet.Cells[1, 2] = "MODEL FILE NAME";
                        xlWorksheet.Cells[1, 3] = "HIERARCHY LEVEL";
                        xlWorksheet.Cells[1, 4] = "CATEGORY";
                        xlWorksheet.Cells[1, 5] = "ELEMENT NAME";
                        xlWorksheet.Cells[1, 6] = "ELEMENT GUID";

                        modelIdx++;

                        //RESET BACK TO FIRST ROW FOR STORING ON NEXT NEW WORKSHEET CREATION
                        rowNum = 2;
                    }
                   
                    //2. RECORD CURRENT EXPORT ITEM VALUES PER COLUMN HEADERS ASSIGNMENT
                    string cellDis = "A" + rowNum.ToString();
                    var rangeDis = xlWorksheet.get_Range(cellDis, cellDis);
                    rangeDis.Value2 = item.ExpDiscipline;

                    string cellModFile = "B" + rowNum.ToString();
                    var rangeModFile = xlWorksheet.get_Range(cellModFile, cellModFile);
                    rangeModFile.Value2 = item.ExpModFile;

                    string cellHiLvl = "C" + rowNum.ToString();
                    var rangeHiLvl = xlWorksheet.get_Range(cellHiLvl, cellHiLvl);
                    rangeHiLvl.Value2 = item.ExpHierLvl;
                    
                    string cellCat = "D" + rowNum.ToString();
                    var rangeCat = xlWorksheet.get_Range(cellCat, cellCat);
                    rangeCat.Value2 = item.ExpCategory;
                    
                    string cellName = "E" + rowNum.ToString();
                    var rangeName = xlWorksheet.get_Range(cellName, cellName);
                    rangeName.Value2 = item.ItemName;

                    string cellId = "F" + rowNum.ToString();
                    var rangeId = xlWorksheet.get_Range(cellId, cellId);
                    rangeId.Value2 = item.ExpGuid;

                    //----------------------------------------------------------------------------------------------

                    //3. SECTION FOR STORING CURRENT EXPORT ITEM PROPERTIES AND VALUES
                    //ITERATE OVER ExportProp and ExportVal LISTS
                    
                    //SETS COLUMN TO START STORING PROPERTY + VALUE INTO CELL
                    colNum = 7;
                    
                    //RETRIEVES CURRENT EXPORT ITEM INDEX NUMBER TO MATCH WITH LIST VALUE IN ItemIdx
                    int indexMatch = ExportProperties.ExportItems.IndexOf(item);
                    var currRange = PropRange(indexMatch); //GOES TO PropRange METHOD TO OBTAIN MINIMUM AND MAXIMUM INDICES
                                                           //OF MATCHING INDEX NUMBER
                    int idxMin = currRange.iMin; //RETURNS MIN. INDEX VALUE OF MATCHED EXPORT ITEM LIST INDEX FROM PropRange
                    int idxMax = currRange.iMax; //RETURNS MAX. INDEX VALUE OF MATCHED EXPORT ITEM LIST INDEX FROM PropRange

                    //USING MIN AND MAX RETURNED VALUES, ITERATES THROUGH ExportProp and ExportVal
                    //TO STORE DATA IN EXCEL FILE PER CURRENT EXPORT ITERATION
                    //ExportProp = Column Header, Export Val = Item Value
                    int i = idxMin;

                    while (i <= idxMax)
                    {

                        // check if column header is empty (unassigned)
                        // create new property column and record value
                        if (xlWorksheet.Cells[1, colNum].Value == null)
                        {
                            xlWorksheet.Cells[1, colNum] = ExportProperties.ExportProp[i];

                            var rangeVal = (Excel.Range)xlWorksheet.Cells[rowNum, colNum];
                            rangeVal.Value2 = ExportProperties.ExportVal[i];
                            i++;
                            colNum = 7;
                        }
                        // check if current property is pointed to same column
                        else if (ExportProperties.ExportProp[i].ToString() == Convert.ToString(xlWorksheet.Cells[1, colNum].Value))
                        {
                            var rangeVal = (Excel.Range)xlWorksheet.Cells[rowNum, colNum];
                            rangeVal.Value2 = ExportProperties.ExportVal[i];
                            i++;
                            colNum = 7;
                        }
                        // if property does not match current column header, 
                        // increment to next column and check if ExportProp matches header
                        // or new header needs to be created
                        else
                        {
                            colNum++;
                        }
                    }

                    /*
                    for (int i = idxMin; i <= idxMax; i++)
                    {
                        var rangeVal = (Excel.Range)xlWorksheet.Cells[rowNum, colNum];
                        rangeVal.Value2 = $"PROPERTY: {ExportProperties.ExportProp[i]}_____VALUE: {ExportProperties.ExportVal[i]}";

                        colNum++;
                    }
                    */
                    //-------------------------------------------------------------------------------------------------------------

                    //SET FOR NEXT EXPORT ITEM 
                    match = false;
                    rowNum++;
                }
                
                //-----------------------------------------------------------------------------------

                //PREPARE DOCUMENT TO PROMPT USER TO SPECIFY FILE SAVE LOCATION

                //GETS CURRENT DATE AND FORMATS FOR DEFAULT FILE NAME
                string exportYr = DateTime.Now.Year.ToString();
                string exportMonth = DateTime.Now.Month.ToString();
                string exportDay = DateTime.Now.Day.ToString();

                if (exportMonth.Length == 1)
                {
                    exportMonth = "0" + exportMonth;
                }

                if (exportDay.Length == 1)
                {
                    exportDay = "0" + exportDay;
                }

                string exportDate = exportYr + exportMonth + exportDay;

                //CREATES NEW INSTANCE FOR WINDOWS EXPLORER SAVE FILE PROMPT
                SaveFileDialog saveModelProperties = new SaveFileDialog();

                //SPECIFIES FILE TYPE EXTENSION TO BE SAVED AS (.XLS)
                saveModelProperties.Title = "Save to...";
                saveModelProperties.Filter = "Excel Workbook | *.xlsx|Excel 97-2003 Workbook | *.xls";
                saveModelProperties.FileName = exportDate + "-System_Property_Data";

                //OPENS SAVE WINDWOS EXPLORER WINDOW FOR USER INPUT
                if (saveModelProperties.ShowDialog() == DialogResult.OK)
                {
                    string path = saveModelProperties.FileName;
                    xlWorkbook.SaveCopyAs(path);
                    xlWorkbook.Saved = true;
                    xlWorkbook.Close(true, Missing.Value, Missing.Value);
                    xlApp.Quit();  
                }

                //WHILE PROCESS IS RUNNING (EXPORT ITEMS --> EXCEL FILE),
                //EXCEL IS INVISIBLE TO USER
                xlApp.Visible = false;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error Writing in Excel File!  Original Message: " + exception.Message);
            }
        }


        //METHOD TO RETURN MINIMUM AND MAXIMUM INDICES IN ExportProp & ExportVal 
        //OF MATCHING CURRENT EXPORT ITEM INDEX NUMBER in ExportItems List
        private static (int iMin, int iMax) PropRange(int indexMatch)
        {
            //CONSTANTS ASSIGNMENT
            int iMin=-1;
            int iMax=-1;
            bool firstMatch = true;

            //ITERATES OVER ItemIdx LIST.
            //IdxCounter KEEPS A RUNNING TOTAL SO DOES NOT HAVE TO START
            //FROM BEGINNING OF 0 BASED LIST...PICKS UP FROM LAST EXPORT ITEM MATCHING INDEX
            for (int i = IdxCounter; i < ExportProperties.ItemIdx.Count; i++)
            {
                    //indexMatch (CURRENT EXPORT ITEM INDEX NUMBER) TO MATCH WITH LIST VALUE IN ItemIdx
                    if (indexMatch == ExportProperties.ItemIdx[i])
                    {
                        if (firstMatch == true)
                        {
                            iMin = i;
                            iMax = i;
                            firstMatch = false;
                        }
                        else
                        {
                            //MAXIMUM VALUE UPDATED TILL NO MORE MATCHING
                            if (i > iMax)
                            {
                                iMax = i;
                                IdxCounter = i;
                            }
                        }
                    }
                }
            
            return (iMin, iMax);
        }


    }
}
