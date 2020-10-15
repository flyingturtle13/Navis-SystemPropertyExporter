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
        //GLOBAL PARAMETERS - ALLOWS ACCESS FROM OTHER CLASSES WITHIN APPLICATION//

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

        //REQUIRES TO BE GLOBAL PARAMETER SINCE A RUNNING COUNT 
        //IS NEEDED AS CYCLES THROUGH ALL MODEL ITEMS FOR EXPORT
        //USED IN CREATING AND ASSOCIATING CATEGORY ExportProp and ExportVal ITEMS
        public static int Idx;

        //CATEGORY PROPERTIES REQUIRE CONTINUED ADDITION FOR ALL EXPORT MODEL ITEMS 
        public static List<string> ExportProp = new List<string>();

        //CATEGORY PROPERTY VALUES REQUIRE CONTINUAL ADDITION FOR ALL EXPORT MODEL ITEMS
        public static List<string> ExportVal = new List<string>();

        //ASSOCIATES CATEGORY & ELEMENT INDEX TO ASSOCIATED
        //ExportProp and ExportVal ITEMS
        public static List<int> ItemIdx = new List<int>();


        //-----------------------------------------------------------------------------------------------------------
        
        //MAIN CLASS METHOD TO PROCESS USER SELECTED MODELS FOR DATA EXPORT
        public static void ProcessModelsSelected()
        {
            foreach (Selected item in UserItems)
            {
                string CurrDis = item.Discipline;
                string CurrModelFile = item.ModFile;
                string UserExportCat = item.SelectCat;
                
                try
                {
                    ModelItem Root = null;

                    //CHECK IF FILE IS NWF
                    foreach (Model model in GetPropertiesModel.DocModel)
                    {
                        if (model.RootItem.DisplayName == item.ModFile)
                        {
                            Root = model.RootItem as ModelItem;
                            ClassTypeCheck_Export(Root, item.HierLvl, CurrDis, CurrModelFile, UserExportCat);
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
                                    ClassTypeCheck_Export(mItem, item.HierLvl, CurrDis, CurrModelFile, UserExportCat);
                                    continue;
                                }
                            }
                        }
                    }
                }
                catch(Exception exception)
                {
                    MessageBox.Show("Error Storing Model Info for Export! Original Message: " + exception.Message);
                }
            }

            //DIRECT TO WriteToExcel CLASS 
            //- EXPORT DESIRED PROPERTIES AND VALUES
            //WriteToExcel.ExcelReport();
            WriteToTxt.txtReport();
        }


        //DETERMINES WHAT HIERARCHY LEVEL TO ACCESS PROPERTIES
        //BASED ON USER INPUT (classType - BUILDING SYSTEM (File), SYSTEM PARTS (Layer), or INDIVIDUAL COMPONENETS (Block)).  
        //DIRECTION FROM GetSystemProperties Method.
        private static void ClassTypeCheck_Export(ModelItem mItem, string classType, string CurrDis, string CurrModelFile, string UserExportCat)
        {
            try
            {
                string CurrExportLvl;
                string CurrExportCat;
                string CurrGuid;
                string CurrEleName;

                //WILL LOOP THROUGH SELECTED FILE AND ALL SUB GEOMTRY ITEMS USING DescendantsAndSelf PROPERTY
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
                                //SETS CURRENT MODEL ITEM REQUIRED PROPERTIES FOR EXPORT
                                //AND PASSES TO NEXT METHOD
                                CurrExportCat = UserExportCat;
                                CurrEleName = subItem1.DisplayName;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_NullCheck(subItem1, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                            }
                            break;

                        case "2": //"Layer"
                            type = "Layer";
                            CurrExportLvl = "Part Types";

                            //CHECK CONDITION IF MODEL AS EXPORTED FROM AUTOCAD
                            //IN THIS CASE, MODEL ITEM IS OF TYPE LAYER
                            if (subItem1.ClassDisplayName == type || subItem1.IsLayer == true)
                            {
                                bool validLayer = false;

                                foreach (ModelItem obj in subItem1.Children)
                                {
                                    if (obj.IsCollection == false)
                                    {
                                        validLayer = true;
                                    }
                                }

                                if (validLayer == true)
                                {
                                    //SETS CURRENT MODEL ITEM REQUIRED PROPERTIES FOR EXPORT
                                    //AND PASSES TO NEXT METHOD
                                    CurrExportCat = UserExportCat;
                                    CurrEleName = subItem1.DisplayName;
                                    CurrGuid = subItem1.InstanceGuid.ToString();
                                    CategoryTypes_NullCheck(subItem1, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                                }
                            }
                            //CHECKS CONDITION WHEN MODEL IS EXPORTED FROM REVIT
                            //IN THIS CASE, REFER TO COLLECTION IF PARENT IS COLLECTION BUT CHILDREN ARE OF DIFFERENT TYPE
                            //(E.G. COMPOSITE, INSERT, GEOMETRY, ETC.)
                            else if (subItem1.IsCollection == true && subItem1.Parent.IsCollection == true)
                            {
                                bool validCollection = false;

                                foreach (ModelItem obj in subItem1.Children)
                                {
                                    if (obj.IsCollection == false)
                                    {
                                        validCollection = true;
                                        
                                    }
                                }

                                if (validCollection == true)
                                {
                                    //SETS CURRENT MODEL ITEM REQUIRED PROPERTIES FOR EXPORT
                                    //AND PASSES TO NEXT METHOD
                                    CurrExportCat = UserExportCat;
                                    CurrEleName = subItem1.DisplayName;
                                    CurrGuid = subItem1.InstanceGuid.ToString();
                                    CategoryTypes_NullCheck(subItem1, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                                }
                            }
                            break;

                        case "3": //Block
                            type = "Block";
                            CurrExportLvl = "Individual Components";

                            //CHECK CONDITION IF MODEL WAS EXPORTED FROM REVIT
                            //IN THIS CASE, MODEL ITEM CLASS TYPE WILL BE BLOCK OR COMPOSITE
                            if ((subItem1.ClassDisplayName == type || subItem1.IsComposite == true) && subItem1.IsInsert == false)
                            {
                                //SETS CURRENT MODEL ITEM REQUIRED PROPERTIES FOR EXPORT
                                //AND PASSES TO NEXT METHOD
                                if (subItem1.DisplayName == "")
                                {
                                    CurrEleName = subItem1.ClassDisplayName;
                                }
                                else
                                {
                                    CurrEleName = subItem1.DisplayName;
                                }
                                CurrExportCat = UserExportCat;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_NullCheck(subItem1, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                            }
                            //CHECK CONDITION IF MODEL WAS EXPORTED FROM AUTOCAD
                            //IN THIS CASE, MODEL ITEM IS OF TYPE GEOMETRY DIRECT SUB TO LAYER SO CHECKS IF PARENT IS LAYER
                            //AND RULES OUT OTHER TYPES.
                            else if (subItem1.Parent.IsLayer == true && subItem1.IsInsert == false && subItem1.IsComposite == false && subItem1.IsCollection == false && subItem1.ClassDisplayName != "Block")
                            {
                                //SETS CURRENT MODEL ITEM REQUIRED PROPERTIES FOR EXPORT
                                //AND PASSES TO NEXT METHOD

                                //MessageBox.Show(subItem1.DisplayName);
                                //MessageBox.Show(subItem1.ClassDisplayName);
                                if (subItem1.DisplayName == "")
                                {
                                    CurrEleName = subItem1.ClassDisplayName;
                                }
                                else
                                {
                                    CurrEleName = subItem1.DisplayName;
                                }
                                CurrExportCat = UserExportCat;
                                CurrGuid = subItem1.InstanceGuid.ToString();
                                CategoryTypes_NullCheck(subItem1, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
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


        //ROUTED TO ACCESS MODEL ITEM ASSOCIATED CATEGORY TYPES AFTER HIEARCHY LEVEL DETAIL(classType) MATCHED
        private static void CategoryTypes_NullCheck(ModelItem subItem, string CurrDis, string CurrModelFile, string CurrExportCat, string CurrEleName, string CurrGuid, string CurrExportLvl)
        {
            Dictionary<string, PropertyCategory> catAvailable = new Dictionary<string, PropertyCategory>();

            //ADD CATEGORIES OF CLASS PropertyCategory AND STRING TO 
            //FACILITATE EASY MATCHING USING STRINGS
            foreach (PropertyCategory oPC in subItem.PropertyCategories)
            {
                catAvailable.Add(oPC.DisplayName, oPC);
            }

            //MessageBox.Show(CurrEleName);
            //foreach (KeyValuePair<string, PropertyCategory> kvp in catAvailable)
            //{
            //    MessageBox.Show(kvp.Key);
            //}

            //CHECK FOR ELEMENT NAME IS BLANK
            if (CurrEleName == "")
            {
                //IF MODEL ITEM DISPLAY NAME IS EMPTY, PASS Item CATEGORY TO ElementNameAssignIfEmpty TO ASSIGN NAME
                if (catAvailable.ContainsKey("Item"))
                {
                    ElementNameAssignIfEmpty(catAvailable["Item"], CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                }
            }
            else
            {
                //IN CASE WHEN ELEMENT NAME IS NOT EMPTY, PROCEED TO CategoryTypes_Export TO RETRIEVE USER SELECTED CATEGORY
                CategoryTypes_Export(catAvailable, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }
        }


        //IF THE MODEL ITEM HAS AN EMPTY ELEMENT NAME, MODEL ITEM CATEGORY Item WILL BE USED TO RETRIEVE
        //Name, Type, OR Layer PROPERTY VALUE TO ASSIGN TO MODEL ITEM ELEMENT NAME.  OTHERWISE, WILL ASSIGN "No Name Assigned"
        private static void ElementNameAssignIfEmpty(PropertyCategory category, string CurrDis, string CurrModelFile, string CurrExportCat, string CurrEleName, string CurrGuid, string CurrExportLvl)
        {
            Dictionary<string, DataProperty> propAvailable = new Dictionary<string, DataProperty>();

            //CREATE DICTIONARY OF CATEGORY PROPERTIES OF CLASS
            //TYPE STRING AND DataProperty <KEY, VALUE>
            foreach (DataProperty oDP in category.Properties)
            {
                propAvailable.Add(oDP.DisplayName, oDP);
            }

            //CHECKS PROPERTY MATCH DESIRED USING CLASS TYPE STRING
            //WHEN MATCH FOUND, ASSIGNS PROPERTY VALUE TO ELEMENT NAME (CurrEleName)
            //AND PROCEEDS TO NEXT METHOD (GetCatProperties_Export)
            if (propAvailable.ContainsKey("Name"))
            {
                CurrEleName = propAvailable["Name"].Value.ToString().Substring(propAvailable["Name"].Value.ToString().IndexOf(':') + 1);
                GetCatProperties_Export(category, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }
            else if (propAvailable.ContainsKey("Type"))
            {
                CurrEleName = propAvailable["Type"].Value.ToString().Substring(propAvailable["Type"].Value.ToString().IndexOf(':') + 1);
                GetCatProperties_Export(category, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }
            else if (propAvailable.ContainsKey("Layer"))
            {
                CurrEleName = propAvailable["Layer"].Value.ToString().Substring(propAvailable["Layer"].Value.ToString().IndexOf(':') + 1);
                GetCatProperties_Export(category, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }
            else
            {
                CurrEleName = "No Name Assigned";
                GetCatProperties_Export(category, CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }

            propAvailable.Clear();
        }


        //ROUTED TO ACCESS MODEL ITEM ASSOCIATED CATEGORY TYPES AFTER HIEARCHY LEVEL DETAIL(classType) MATCHED
        private static void CategoryTypes_Export(Dictionary<string, PropertyCategory> catAvailable, string CurrDis, string CurrModelFile, string CurrExportCat, string CurrEleName, string CurrGuid, string CurrExportLvl)
        {
            //Dictionary<string, PropertyCategory> catAvailable = new Dictionary<string, PropertyCategory>();

            //CREATE DICTIONARY OF CATEGORY OF CLASS
            //TYPE STRING AND PropertyCategory <KEY, VALUE>
            //foreach (PropertyCategory oPC in subItem.PropertyCategories)
            //{
            //    catAvailable.Add(oPC.DisplayName, oPC);
            //    MessageBox.Show(oPC.DisplayName);
            //}

            //foreach (KeyValuePair<string, PropertyCategory> kvp in catAvailable)
            //{
            //    MessageBox.Show(kvp.Key);
            //}

            //2. SEARCH IF DESIRED USER CATEGORY EXISTS FOR MODEL ITEM
            //PROCEEDS TO OBTAIN PROPERTIES OF CATEGORY
            //SPECIFIED BY USER IF MATCH FOUND
            //MessageBox.Show(CurrExportCat.ToString());
            if (catAvailable.ContainsKey(CurrExportCat))
            {
                //MessageBox.Show("Found");

                //IF MATCH CATEGORY MATCH FOUND OF CLASS TYPE PropertyCategory, PROCEEDS TO GetProerties_Export METHOD
                GetCatProperties_Export(catAvailable[CurrExportCat], CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                catAvailable.Clear();
            }
            //IF CATEGORY DOESN'T EXIST FOR MODEL ITEM, DEFAULT TO Item (TYPICALLY EXISTS FOR ALL MODEL ITEMS)
            else if (catAvailable.ContainsKey("Item"))
            {
                //ELEMENT NAME REDIRECT FOR ASSIGNMENT IF INITIAL
                //CATEGORY 'DisplayName' == null
                CurrExportCat = "Item";
                    
                //PROCEED TO GetCatProperties_Export METHOD
                GetCatProperties_Export(catAvailable[CurrExportCat], CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
                catAvailable.Clear();
            }
            else
            {
                //IF MODEL ITEM HAS NOT CATEGORIES DEFAULT TO "null"
                //FOR PROPERTY AND VALUE AND PROCEEDS TO ExportItemSet
                ItemIdx.Add(Idx);
                ExportProp.Add("null");
                ExportVal.Add("null");
                catAvailable.Clear();

                ExportItemsSet(CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
            }
        }


        //ACCESS AVAILABLE PROPERTIES AND VALUES PER CATEGORY SELECTED BY USER
        private static void GetCatProperties_Export(PropertyCategory category, string CurrDis, string CurrModelFile, string CurrExportCat, string CurrEleName, string CurrGuid, string CurrExportLvl)
        {
            //UPDATES 3 LISTS: Idx, ExportProp, and ExportVal
            
            //CHECKS PROPERTIES EXIST FOR THE CATEGORY
            if (category.Properties.Count > 0)
            { 
                foreach (DataProperty oDP in category.Properties)
                {
                    //CHECKS FOR PROPERTY VALUE IS NULL OR EMPTY
                    //IF NULL OR EMPTY, ASSIGNS "null" FOR VALUE
                    if (oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':') + 1) == "" || oDP.Value == null)
                    {
                        ItemIdx.Add(Idx);
                        ExportProp.Add(oDP.DisplayName);
                        ExportVal.Add("null");
                    }
                    //OTHERWISE, ADDS MODEL ITEM VALUES TO LISTS (ItemIdx (ASSOCIATES
                    //PROPERTIES AND VALUES TO CORRECT EXPORTED MODEL ITEM DUE TO NUMBER OF POPERTIES
                    //VARYING PER CATEGORY AND MODEL ITEM), ExportProp, ExportVal)
                    else
                    {
                        ItemIdx.Add(Idx);
                        ExportProp.Add(oDP.DisplayName);
                        //ISSUES WITH ToDisplayString() IN AUTODESK API.  Using Substring() and IndexOf() METHODS
                        //TO REMOVE UNWANTED CHARACTERS IN STRING
                        ExportVal.Add(oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':') + 1));
                    }
                }
            }
            //IF THERE ARE NO PROPERTIES IN THE CATEGORY, THEN ASSIGNS "null"
            //FOR BOTH ExportProp AND ExportVal.
            else
            {
                ItemIdx.Add(Idx);
                ExportProp.Add("null");
                ExportVal.Add("null");
            }

            //PROCEEDS TO ExportItemsSet METHOD TO ASSIGN VALUES TO ExportItems (RECORDS TO EXCEL FILE)
            ExportItemsSet(CurrDis, CurrModelFile, CurrExportCat, CurrEleName, CurrGuid, CurrExportLvl);
        }
        

        private static void ExportItemsSet(string CurrDis, string CurrModelFile, string CurrExportCat, string CurrEleName, string CurrGuid, string CurrExportLvl)
        {
            //STORE VALUES IN ExportItems WHICH IS USED TO RECORD TO EXCEL FILE
            ExportItems.Add(new Export
            {
                ExpDiscipline = CurrDis,
                ExpModFile = CurrModelFile,
                ExpHierLvl = CurrExportLvl,
                ExpCategory = CurrExportCat,
                ItemName = CurrEleName,
                ExpGuid = CurrGuid,
            });

            //INCREMENT Idx (GLOBAL PARAMETER) FOR NEXT MODEL ITEM
            //TO ASSOCIATE TO ITS PROPERTIES (ExportProp) AND VALUES (ExportVal)
            Idx++;
        }

    }
   

    //CLASS TO BIND USER SELECTED PARAMETERS TO COLUMNS IN MODELSSELECTED_LISTVIEW
    //USING UserItems Observable Collection
    public class Selected
    {
        public string Discipline { get; set; }
        public string ModFile { get; set; }
        public string HierLvl { get; set; }
        public string SelectCat { get; set; }
    }


    //CLASS BINDING CONTAINER FOR ExportItems Observable Collection FOR
    //RECORDING TO EXCEL FILE.
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
