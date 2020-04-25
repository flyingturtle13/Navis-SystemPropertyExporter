using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    class ExportProperties
    {
        private static ObservableCollection<Selected> _userItems;
        
        public static ObservableCollection<Selected> UserItems
        {
            get
            {
                if (_userItems == null)
                {
                    _userItems = new ObservableCollection<Selected>();
                }
                return _userItems;
            }
            set
            {
                _userItems = value;
            }
        }

        private static ObservableCollection<Export> _exportItems;

        public static ObservableCollection<Export> ExportItems
        {
            get
            {
                if (_exportItems == null)
                {
                    _exportItems = new ObservableCollection<Export>();
                }
                return _exportItems;
            }
            set
            {
                _exportItems = value;
            }
        }

        public static int Selected_HierLvl { get; set; }

        public static string Selected_Cat { get; set; }

        public static ModelItem Root { get; set; }

        public static int Idx;

        public static List<string> ExportProp = new List<string>();

        public static List<string> ExportVal = new List<string>();

        public static List<int> ItemIdx = new List<int>();

        public static string CurrDis { get; set; }

        public static string CurrModelFile { get; set; }

        public static string CurrExportLvl { get; set; }

        public static string CurrExportCat { get; set; }

        public static string CurrEleName { get; set; }

        public static string CurrGuid { get; set; }

        public static string UserExportCat { get; set; }

        //public static int Cnt;


        //-----------------------------------------------------------------------
        

        public static void ProcessModelsSelected()
        {
            //ExportPhase = true;

            foreach (Selected item in UserItems)
            {
                CurrDis = item.Discipline;
                CurrModelFile = item.ModFile;
                UserExportCat = item.SelectCat;
                
                try
                {
                    //CHECK IF FILE IS NWF
                    foreach (Model model in GetPropertiesModel.DocModel)
                    {
                        if (model.RootItem.DisplayName == item.ModFile)
                        {
                            Root = model.RootItem as ModelItem;
                            ClassTypeCheck_Export(Root, item.HierLvl);
                            break;
                        }
                    }


                    //ENTERS IF FILE IS NWD (GO NEXT LEVEL TO SEARCH FOR MODEL FILES)
                    if (Root == null)
                    {
                        foreach (Model model in GetPropertiesModel.DocModel)
                        {
                            ModelItem root = model.RootItem as ModelItem;
                            foreach (ModelItem mItem in root.Children)
                            {
                                if (mItem.DisplayName == item.ModFile)
                                {
                                    ClassTypeCheck_Export(mItem, item.HierLvl);
                                    continue;
                                }
                            }
                        }
                    }
                }
                catch(Exception exception)
                {
                    MessageBox.Show("Error! Original Message: " + exception.Message);
                }
            }

            WriteToExcel.ExcelReport();
        }


        //DETERMINES WHAT HIERARCHY LEVEL TO ACCESS PROPERTIES
        //BASED ON USER INPUT (classType - File, Layer, or Block).  DIRECTION FROM GetSystemProperties Method.
        private static void ClassTypeCheck_Export(ModelItem mItem, string classType)
        {
            try
            {
                foreach (ModelItem subItem1 in mItem.DescendantsAndSelf)
                {
                    string type = "";

                    switch (classType)
                    {
                        case "1": //"File"
                            type = "File";
                            CurrExportLvl = "Overall System";

                            if (subItem1.ClassDisplayName == type)
                            {
                                CurrExportCat = UserExportCat;
                                CurrEleName = subItem1.DisplayName;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_Export(subItem1);
                            }
                            break;

                        case "2": //"Layer"
                            type = "Layer";
                            CurrExportLvl = "Part Types";

                            if (subItem1.ClassDisplayName == type || subItem1.IsLayer == true)
                            {
                                //foreach (ModelItem obj in subItem1.Children)
                                //{
                                //    if (obj.IsInsert == false && obj.IsComposite == false && obj.IsCollection == false && obj.ClassDisplayName != "Block")
                                //    {
                                //        continue;
                                //    }
                                //    else
                                //    {
                                        //MessageBox.Show(subItem1.ClassDisplayName + ", " + subItem1.DisplayName);
                                        CurrExportCat = UserExportCat;
                                        CurrEleName = subItem1.DisplayName;
                                        CurrGuid = subItem1.InstanceGuid.ToString();
                                        CategoryTypes_Export(subItem1);
                                //    }
                                //}
                            }
                            break;

                        case "3": //Block
                            type = "Block";
                            CurrExportLvl = "Individual Components";
                            
                            if ((subItem1.ClassDisplayName == type || subItem1.IsComposite == true) && subItem1.IsInsert == false)
                            {
                                //MessageBox.Show(subItem1.ClassDisplayName + ", " + subItem1.DisplayName);
                                CurrExportCat = UserExportCat;
                                CurrEleName = subItem1.DisplayName;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_Export(subItem1);
                            }
                            else if (subItem1.Parent.IsLayer == true && subItem1.IsInsert == false && subItem1.IsComposite == false && subItem1.IsCollection == false && subItem1.ClassDisplayName != "Block")
                            {
                                CurrExportCat = UserExportCat;
                                CurrEleName = subItem1.DisplayName;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_Export(subItem1);
                            }
                            break;
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        //ROUTED TO ACCESS MODEL ITEM ASSOCIATED CATEGORY TYPES AFTER HIEARCHY LEVEL DETAIL(classType)
        //MATCHED
        private static void CategoryTypes_Export(ModelItem subItem)
        {
            Dictionary<string, PropertyCategory> catAvailable = new Dictionary<string, PropertyCategory>();
            
            try
            {
                foreach (PropertyCategory oPC in subItem.PropertyCategories)
                {
                    catAvailable.Add(oPC.DisplayName, oPC);
                }

                //1. CHECK FOR ELEMENT NAME IS BLANK
                if (CurrEleName == "")
                {
                    if (catAvailable.ContainsKey("Item"))
                    {
                        ElementNameAssignIfEmpty(catAvailable["Item"]);
                    }
                }

                //2. SEARCH IF DESIRED USER CATEGORY EXISTS FOR MODEL ITEM
                //PROCEEDS TO OBTAIN PROPERTIES OF CATEGORY
                //SPECIFIED BY USER IF MATCH FOUND
                //MessageBox.Show($"{UserExportCat} ---- {CurrExportCat} ---- {catAvailable.ContainsKey(CurrExportCat)}");
                if (catAvailable.ContainsKey(CurrExportCat))
                {
                    GetCatProperties_Export(catAvailable[CurrExportCat]);
                    catAvailable.Clear();
                    
                }
                else if (catAvailable.ContainsKey("Item"))
                {
                    //MessageBox.Show("chec 2");
                    //ELEMENT NAME REDIRECT FOR ASSIGNMENT IF INITIAL
                    //ELEMENT NAME FROM 'DisplayName' == null
                    CurrExportCat = "Item";
                    
                    GetCatProperties_Export(catAvailable[CurrExportCat]);
                    catAvailable.Clear();
                }
                else
                {
                    ItemIdx.Add(Idx);
                    ExportProp.Add("null");
                    ExportVal.Add("null");
                    catAvailable.Clear();

                    ExportItemsSet();
                }
                
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        private static void ElementNameAssignIfEmpty(PropertyCategory category)
        {
            //bool match = false;
            Dictionary<string, DataProperty> propAvailable = new Dictionary<string, DataProperty>();


            foreach (DataProperty oDP in category.Properties)
            {
                propAvailable.Add(oDP.DisplayName, oDP);
            }

            if (propAvailable.ContainsKey("Name"))
            {
                //MessageBox.Show($"Name - {catAvailable["Name"].Value.ToString().Substring(catAvailable["Name"].Value.ToString().IndexOf(':') + 1)}");
                CurrEleName = propAvailable["Name"].Value.ToString().Substring(propAvailable["Name"].Value.ToString().IndexOf(':') + 1);
            }
            else if (propAvailable.ContainsKey("Type"))
            {
                //MessageBox.Show($"Type - {catAvailable["Type"].Value.ToString().Substring(catAvailable["Type"].Value.ToString().IndexOf(':') + 1)}");
                CurrEleName = propAvailable["Type"].Value.ToString().Substring(propAvailable["Type"].Value.ToString().IndexOf(':') + 1);
            }
            else if (propAvailable.ContainsKey("Layer"))
            {
                //MessageBox.Show($"Layer - {catAvailable["Layer"].Value.ToString().Substring(catAvailable["Layer"].Value.ToString().IndexOf(':') + 1)}");
                CurrEleName = propAvailable["Layer"].Value.ToString().Substring(propAvailable["Layer"].Value.ToString().IndexOf(':') + 1);
            }
            else
            {
                CurrEleName = "No Name Assigned";
            }

            propAvailable.Clear();
        }
        ////IN THE CASE WHEN CATEGORY DESIRED DOES NOT EXIST IN CURRENT ELEMENT (SOME ELEMENTS DIFFER IN CATEGOORIES)
        ////THIS STORES NULL FOR PROPERTIES FOR THE ELEMENT WHEN TYPE OF CATEGORY DOESN'T EXIST.
        ////Idx LIST FOR ACCOUNTING FOR PROPERTIES ASSOCIATED TO CATEGORY GETS UPDATED TO KEEP CATEGORIES TO PROPERTIES LINKED.
        //private static  void CatProperties_NoMatch()
        //{
        //    ItemIdx.Add(Idx);
        //    ExportProp.Add("null");
        //    ExportVal.Add("null");

        //    Idx++;
        //}


        //ACCESS AVAILABLE PROPERTIES PER CATEGORY SELECTED BY USER
        private static void GetCatProperties_Export(PropertyCategory category)
        {
            //UPDATE 3 LISTS: Idx, ExportProp, and ExportVal
            //
            try
            {
                if (category.Properties.Count > 0)
                { 
                    foreach (DataProperty oDP in category.Properties)
                    {
                        if (oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':') + 1) == "" || oDP.Value == null)
                        {
                            ItemIdx.Add(Idx);
                            ExportProp.Add(oDP.DisplayName);
                            ExportVal.Add("null");
                        }
                        else
                        {
                            ItemIdx.Add(Idx);
                            ExportProp.Add(oDP.DisplayName);
                            ExportVal.Add(oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':') + 1));
                        }
                    }
                }
                else
                {
                    ItemIdx.Add(Idx);
                    ExportProp.Add("null");
                    ExportVal.Add("null");
                }

                ExportItemsSet();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
           
        }
        

        private static void ExportItemsSet()
        {
            //MessageBox.Show(CurrExportCat + "--" + CurrEleName);
            //STORE VALUES
            ExportItems.Add(new Export
            {
                ExpDiscipline = CurrDis,
                ExpModFile = CurrModelFile,
                ExpHierLvl = CurrExportLvl,
                ExpCategory = CurrExportCat,
                ItemName = CurrEleName,
                ExpGuid = CurrGuid,
            });

            Idx++;
        }

    }
   

    //CLASS TO BIND USER SELECTED PARAMETERS TO COLUMNS IN MODELSSELECTED_LISTVIEW
    public class Selected
    {
        public string Discipline { get; set; }
        public string ModFile { get; set; }
        public string HierLvl { get; set; }
        public string SelectCat { get; set; }
    }


    public class Export
    {
        public string ExpDiscipline { get; set; }
        public string ExpModFile { get; set; }
        public string ExpHierLvl { get; set; }
        public string ExpCategory { get; set; }
        public string ItemName { get; set; }
        public string ExpGuid { get; set; }

    }
}
