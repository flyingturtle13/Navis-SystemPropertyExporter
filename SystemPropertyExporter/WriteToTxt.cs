using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Microsoft.Win32;
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
    class WriteToTxt
    {
        //REQUIRED TO BE GLOBAL PARAMETER SO A RUNNING TOTAL CAN BE KEPT
        //PREVENTS HAVING TO CYCLE THROUGH ENTIRE 0 BASED ItemIdx List in RangeProp METHOD
        public static int TxtIdxCounter { get; set; }


        public static void txtReport()
        {

            List<string> header = new List<string>();
            List<string> data = new List<string>();

            try
            {
                TxtIdxCounter = 0;
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

                //CREATES NEW VARIABLE INSTANCE - ALLOWS TO OPEN WINDOWS EXPLORER SAVE FILE PROMPT
                string filename = "";
                SaveFileDialog saveExportData = new SaveFileDialog();

                //SETS FILE TYPE TO BE SAVED AS .TXT
                saveExportData.Title = "Save to...";
                saveExportData.Filter = "Text Documents | *.txt";
                saveExportData.FileName = exportDate + "-System_Property_Data";

                //OPENS WINDOWS EXPLORER TO BEGIN LIST SAVE PROCESS
                if (saveExportData.ShowDialog() == true)
                {
                    filename = saveExportData.FileName.ToString();
                    string savePath = $"{Path.GetDirectoryName(filename)}\\";
                   
                    //CHECKS USER HAS INPUTED A NAME FOR THE FILE
                    if (filename != "")
                    {
                        using (StreamWriter sw = new StreamWriter(filename))
                        {
                            //POPULATES TXT FILE WITH LIST ITEMS SEPARATED BY "--" USING StreamWriter & CLOSES WHEN COMPLETE
                            header.Add("Discipline");
                            header.Add("Model File Name");
                            header.Add("Hierarchy Level");
                            header.Add("Category");
                            header.Add("Elemenet Name");
                            header.Add("Element GUID");

                            foreach (Export item in ExportProperties.ExportItems)
                            {
                                data.Add(item.ExpDiscipline);
                                data.Add(item.ExpModFile);
                                data.Add(item.ExpHierLvl);
                                data.Add(item.ExpCategory);
                                data.Add(item.ItemName);
                                data.Add(item.ExpGuid);

                                //RETRIEVES CURRENT EXPORT ITEM INDEX NUMBER TO MATCH WITH LIST VALUE IN ItemIdx
                                int indexMatch = ExportProperties.ExportItems.IndexOf(item);
                                var currRange = PropRange(indexMatch); //GOES TO PropRange METHOD TO OBTAIN MINIMUM AND MAXIMUM INDICES
                                                                       //OF MATCHING INDEX NUMBER
                                int idxMin = currRange.iMin; //RETURNS MIN. INDEX VALUE OF MATCHED EXPORT ITEM LIST INDEX FROM PropRange
                                int idxMax = currRange.iMax; //RETURNS MAX. INDEX VALUE OF MATCHED EXPORT ITEM LIST INDEX FROM PropRange

                                int i = idxMin;
                                int colNum = 6;
                                //MessageBox.Show(i.ToString());
                                while (i <= idxMax)
                                {
                                    // check if column header is empty (unassigned)
                                    // create new property column and record value
                                    if (colNum >= header.Count)
                                    {
                                        header.Add(ExportProperties.ExportProp[i].ToString());
                                        data.Add(ExportProperties.ExportVal[i].ToString());

                                        i++;
                                        colNum = 6;
                                    }
                                    // check if current property is pointed to same column
                                    else if (ExportProperties.ExportProp[i].ToString() == header[colNum].ToString())
                                    {
                                        if (colNum >= data.Count)
                                        {
                                            data.Add(ExportProperties.ExportVal[i].ToString());
                                        }
                                        else
                                        {
                                            data[colNum] = ExportProperties.ExportVal[i].ToString();
                                        }

                                        i++;
                                        colNum = 6;
                                    }
                                    // if property does not match current column header, 
                                    // increment to next column and check if ExportProp matches header
                                    // or new header needs to be created
                                    else if (data.Count < header.Count)
                                    {
                                        data.Add("null");
                                        colNum++;
                                    }
                                    else
                                    {
                                        colNum++;
                                    }

                                }

                                // Write values to text file
                                foreach (string value in data)
                                {
                                    if (data.IndexOf(value) == 0)
                                    {
                                        sw.Write(value);
                                    }
                                    else
                                    {
                                        sw.Write("^" + value);
                                    }
                                }

                                sw.WriteLine("");
                                data.Clear();
                            }
                            sw.Dispose();
                            sw.Close();

                        }

                        
                        //SETS FILE TYPE TO BE SAVED AS .TXT FOR COLUMN HEADERS
                        filename = exportDate + "-System_Property_Headers.txt";

                        using (StreamWriter swHeader = new StreamWriter(savePath + filename))
                        {
                            foreach (string title in header)
                            {
                                if (header.IndexOf(title) == 0)
                                {
                                    swHeader.Write(title.ToString());
                                }
                                else
                                {
                                    swHeader.Write("^" + title.ToString());
                                }
                            }

                            swHeader.Dispose();
                            swHeader.Close();
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error Writing in Txt File!  Original Message: " + exception.Message);
            }
        }


        //METHOD TO RETURN MINIMUM AND MAXIMUM INDICES IN ExportProp & ExportVal 
        //OF MATCHING CURRENT EXPORT ITEM INDEX NUMBER in ExportItems List
        private static (int iMin, int iMax) PropRange(int indexMatch)
        {
            //CONSTANTS ASSIGNMENT
            int iMin = -1;
            int iMax = -1;
            bool firstMatch = true;

            //ITERATES OVER ItemIdx LIST.
            //IdxCounter KEEPS A RUNNING TOTAL SO DOES NOT HAVE TO START
            //FROM BEGINNING OF 0 BASED LIST...PICKS UP FROM LAST EXPORT ITEM MATCHING INDEX
            for (int i = TxtIdxCounter; i < ExportProperties.ItemIdx.Count; i++)
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
                            TxtIdxCounter = i;
                        }
                    }
                }
            }

            return (iMin, iMax);
        }
    }
}
