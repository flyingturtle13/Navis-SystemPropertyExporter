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

        public static int Idx = 0;

        public static List<string> ExportProp = new List<string>();

        public static List<string> ExportVal = new List<string>();

        public static List<int> ItemIdx = new List<int>();

        public static string CurrDis { get; set; }

        public static string CurrModelFile { get; set; }

        public static string CurrExportLvl { get; set; }

        public static string CurrExportCat { get; set; }


        //-----------------------------------------------------------------------


        public static void ProcessModelsSelected()
        {
            //ExportPhase = true;

            foreach (Selected item in UserItems)
            {
                CurrDis = item.Discipline;
                CurrModelFile = item.ModFile;
                CurrExportCat = item.SelectCat;
                

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
                                CategoryTypes_Export(subItem1);
                            }
                            break;

                        case "2": //"Layer"
                            type = "Layer";
                            CurrExportLvl = "Part Types";

                            if (subItem1.ClassDisplayName == type || subItem1.IsLayer == true)
                            {
                                foreach (ModelItem obj in subItem1.Children)
                                {
                                    if (obj.IsInsert == false && obj.IsComposite == false && obj.IsCollection == false && obj.ClassDisplayName != "Block")
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        //MessageBox.Show(subItem1.ClassDisplayName + ", " + subItem1.DisplayName);
                                        CategoryTypes_Export(subItem1);
                                    }
                                }
                            }
                            break;

                        case "3": //Block
                            type = "Block";
                            CurrExportLvl = "Individual Components";

                            if (subItem1.ClassDisplayName == type || subItem1.IsComposite == true)
                            {

                                //MessageBox.Show(subItem1.ClassDisplayName + ", " + subItem1.DisplayName);
                                CategoryTypes_Export(subItem1);
                            }
                            else if (subItem1.IsLayer == true)
                            {
                                foreach (ModelItem obj in subItem1.Children)
                                {
                                    if (obj.IsInsert == false && obj.IsComposite == false && obj.IsCollection == false && obj.ClassDisplayName != "Block")
                                    {
                                        CategoryTypes_Export(subItem1);
                                    }
                                }
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
            try
            {
                ExportItems.Add(new Export
                {
                    ExpDiscipline = CurrDis,
                    ExpModFile = CurrModelFile,
                    ExpHierLvl = CurrExportLvl,
                    ExpCategory = CurrExportCat,
                    ItemName = subItem.DisplayName
                });

                foreach (PropertyCategory oPC in subItem.PropertyCategories)
                {
                    if (oPC.DisplayName == CurrExportCat)
                    {
                        GetCatProperties_Export(oPC);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }

        //ACCESS AVAILABLE PROPERTIES PER CATEGORY SELECTED BY USER
        public static void GetCatProperties_Export(PropertyCategory category)
        {
            try
            {
                if (category.Properties.Count > 0)
                {
                    foreach (DataProperty oDP in category.Properties)
                    {
                        ItemIdx.Add(Idx);
                        ExportProp.Add(oDP.DisplayName);
                        ExportVal.Add(oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':') + 1));

                    }
                }
                else
                {
                    ItemIdx.Add(Idx);
                    ExportProp.Add("null");
                    ExportVal.Add("null");
                }

                Idx++;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
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
    }
}
