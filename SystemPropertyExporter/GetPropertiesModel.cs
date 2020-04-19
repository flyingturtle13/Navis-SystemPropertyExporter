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

namespace SystemPropertyExporter
{
    class GetPropertiesModel
    {
        public static DocumentModels DocModel { get; set; }

        public static List<PropertyCategory> CurrCategories = new List<PropertyCategory>();
        
        private static ObservableCollection<string> _modelList;

        public static ObservableCollection<string> ModelList
        {
            get
            {
                if (_modelList == null)
                {
                    _modelList = new ObservableCollection<string>();
                }
                return _modelList;
            }
            set
            {
                _modelList = value;
            }
        }

        private static ObservableCollection<Property> _returnProp;

        public static ObservableCollection<Property> ReturnProp
        {
            get
            {
                if(_returnProp == null)
                {
                    _returnProp = new ObservableCollection<Property>();
                }
                return _returnProp;
            }
            set
            {
                _returnProp = value;
            }
        }

        private static ObservableCollection<Category> _returnCategories;

        public static ObservableCollection<Category> ReturnCategories
        {
            get
            {
                if (_returnCategories == null)
                {
                    _returnCategories = new ObservableCollection<Category>();
                }
                return _returnCategories;
            }
            set
            {
                _returnCategories = value;
            }
        }
        
        public static ModelItem Root { get; set; }
        

        //---------------------------------------------------------------------------------------


        //STEP 1
        //THIS METHOD TAKES INPUTS FROM UserInput FORM.
        //DETERMINES IF CURRENT PROJECT IS NWF (LIVE) OR IF NWD (SNAPSHOT)
        public static void GetSystemProperties(string displayName, string classType)
        {
            ReturnProp.Clear();
            CurrCategories.Clear();
            ReturnCategories.Clear();
            
            //CHECK IF FILE IS NWF
            foreach (Model model in DocModel)
            {
                if (model.RootItem.DisplayName == displayName)
                {
                    Root = model.RootItem as ModelItem;
                    ClassTypeCheck(Root, classType);
                    break;
                }
            }

            //ENTERS IF FILE IS NWD (GO NEXT LEVEL TO SEARCH FOR MODEL FILES)
            if (Root == null)
            {
                foreach (Model model in DocModel)
                {
                    ModelItem root = model.RootItem as ModelItem;
                    foreach (ModelItem item in root.Children)
                    {
                        if (item.DisplayName == displayName)
                        {
                            ClassTypeCheck(item, classType);
                            continue;
                        }
                    }
                }
            }
        }


        //DETERMINES WHAT HIERARCHY LEVEL TO ACCESS PROPERTIES
        //BASED ON USER INPUT (classType - File, Layer, or Block).  DIRECTION FROM GetSystemProperties Method.
        private static int ClassTypeCheck(ModelItem item, string classType) {

            foreach (ModelItem subItem1 in item.DescendantsAndSelf)
            {
                switch (classType)
                {
                    case "File":
                       if (subItem1.ClassDisplayName == classType)
                       {
                            CategoryTypes(subItem1);
                            return 0;
                       }
                       break;

                    case "Layer":
                        if (subItem1.ClassDisplayName == classType || subItem1.IsLayer == true)
                        {
                            CategoryTypes(subItem1);
                            return 0;
                        }
                        break;

                    case "Block":
                        if (subItem1.ClassDisplayName == classType || subItem1.IsComposite == true)
                        {
                             CategoryTypes(subItem1);
                             return 0;
                        }
                        break;
                }
            }
            return 0;
        }


        //ROUTED TO ACCESS MODEL ITEM ASSOCIATED CATEGORY TYPES AFTER HIEARCHY LEVEL DETAIL(classType)
        //MATCHED
        private static void CategoryTypes(ModelItem item)
        {
            foreach (PropertyCategory oPC in item.PropertyCategories)
            {
                //STORES IN ReturnCategories TO DISPLAY AVAILABLE CATEGORIES IN UserInput FORM IN CatProp_ListView
                //CurrCategories STORES CATEGORIES AS PropertyCategory (Navis API) TYPE
                //THIS WILL BE ACCESSED IN STEP 2 (GetCatProperties()) AFTER USER HAS SELECTED WHICH CATEGORY TO ACCESS
                CurrCategories.Add(oPC);
                ReturnCategories.Add(new Category
                {
                    CatName = oPC.DisplayName
                });
            }
        }


        //STEP 2
        //ROUTED HERE FROM UserInput FORM AFTER USER SELECTS CATEGORY TYPE TO ACCESS ASSOCIATED PROPERTIES 
        //SELECTED CATEGORY IS PASSED AS CatNameSelected
        public static void GetCatProperties(string CatNameSelected)
        {
            foreach (PropertyCategory category in CurrCategories)
            {

                if (category.DisplayName == CatNameSelected)
                {
                    foreach (DataProperty oDP in category.Properties)
                    {
                        //STORES IN ReturnProp TO BE DISPLAYED IN UserInput FORM IN Prop_ListView
                        ReturnProp.Add(new Property
                        {
                            PropName = oDP.DisplayName,
                            //ISSUES WITH ToDisplayString() IN AUTODESK API.  Using Substring() and IndexOf() METHODS
                            //TO REMOVE UNWANTED CHARACTERS IN STRING
                            ValEx = oDP.Value.ToString().Substring(oDP.Value.ToString().IndexOf(':')+1)
                        });
                    }
                }
            }
        }


        //ROUTED FROM StarMain TO STORE PROJECT MODELS IN modelList LIST 
        //TO BE DISPLAYED IN UserInput USING Models_ComboBox
        public static void GetCurrModels()
        {
            foreach (Model model in DocModel)
            {
                ModelList.Add(model.RootItem.DisplayName);
            }

            foreach (Model model in DocModel)
            {
                ModelItem root = model.RootItem as ModelItem;
                foreach (ModelItem item in root.Children)
                {
                    ModelList.Add(item.DisplayName);
                }
            }
        }
    }
    

    public class Category
    {
        public string CatName { get; set; }
    }


    public class Property
    {
        public string PropName { get; set; }
        public string ValEx { get; set; }
    }
}
