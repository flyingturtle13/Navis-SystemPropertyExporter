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
        //GLOBAL PARAMETERS - ALLOWS ACCESS FROM OTHER CLASSES WITHIN APPLICATION//
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

        //REQUIRES TO BE GLOBAL SINCE A RUNNING COLLECTION IS NEEDED AS CYCLES THROUGH ALL MODEL ITEMS
        public static List<string> catDuplicate = new List<string>();
        

        //---------------------------------------------------------------------------------------


        //STEP 1
        //THIS METHOD TAKES INPUTS FROM UserInput FORM.
        //DETERMINES IF CURRENT PROJECT IS NWF (LIVE) OR IF NWD (SNAPSHOT)
        public static void GetSystemProperties(string displayName, string classType)
        {
            ReturnProp.Clear();
            CurrCategories.Clear();
            ReturnCategories.Clear();
            catDuplicate.Clear();
            ModelItem Root = null;

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
                        }
                    }
                }
            }
        }


        //DETERMINES WHAT HIERARCHY LEVEL TO ACCESS PROPERTIES
        //BASED ON USER INPUT (classType - BUILDING SYSTEM (File), SYSTEM PARTS (Layer), or INDIVIDUAL COMPONENETS (Block)).  
        //DIRECTION FROM GetSystemProperties Method.
        private static void ClassTypeCheck(ModelItem item, string classType) {

            //WILL LOOP THROUGH SELECTED FILE AND ALL SUB GEOMTRY ITEMS USING DescendantsAndSelf PROPERTY
            foreach (ModelItem subItem1 in item.DescendantsAndSelf)
            {
                switch (classType)
                {
                    case "File":
                       if (subItem1.ClassDisplayName == classType)
                       {
                            CategoryTypes(subItem1);
                       }
                       break;

                    case "Layer":
                        //CHECK CONDITION IF MODEL AS EXPORTED FROM AUTOCAD
                        //IN THIS CASE, MODEL ITEM IS OF TYPE LAYER
                        if (subItem1.ClassDisplayName == classType || subItem1.IsLayer == true)
                        {
                            bool validLayer = false;

                            foreach(ModelItem obj in subItem1.Children)
                            {
                                if (obj.IsCollection == false)
                                {
                                    validLayer = true;
                                    break;
                                }
                            }

                            if (validLayer == true)
                            {
                                CategoryTypes(subItem1);
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
                                    break;
                                }
                            }

                            if (validCollection == true)
                            {
                                CategoryTypes(subItem1);
                            }
                        }
                        break;

                    case "Block":

                        //CHECK CONDITION IF MODEL WAS EXPORTED FROM REVIT
                        //IN THIS CASE, MODEL ITEM CLASS TYPE WILL BE BLOCK OR COMPOSITE
                        if (subItem1.ClassDisplayName == classType || subItem1.IsComposite == true)
                        {
                             CategoryTypes(subItem1);
                        }
                        //CHECK CONDITION IF MODEL WAS EXPORTED FROM AUTOCAD
                        //IN THIS CASE, MODEL ITEM IS OF TYPE GEOMETRY DIRECT SUB TO LAYER SO CHECKS IF PARENT IS LAYER
                        //AND RULES OUT OTHER TYPES.
                        else if (subItem1.Parent.IsLayer == true && subItem1.IsInsert == false && subItem1.IsComposite == false && subItem1.IsCollection == false && subItem1.ClassDisplayName != "Block")
                        {
                            CategoryTypes(subItem1);
                        }
                        break;
                }
            }
        }


        //ROUTED TO ACCESS MODEL ITEM ASSOCIATED CATEGORY TYPES AFTER HIEARCHY LEVEL DETAIL(classType) MATCHED
        private static void CategoryTypes(ModelItem item)
        {
            //CYCLES THROUGH ALL AVAILABLE CATEGORIES PER MODEL ITEM
            foreach (PropertyCategory oPC in item.PropertyCategories)
            {
                //MAKES A SINGLE COLLECTION OF AVAILABLE CATEGORIES
                //CHECKS TO PREVENT DUPLICATES OF CATEGORIES
                if (!catDuplicate.Contains(oPC.DisplayName))
                {   
                    //STORES IN ReturnCategories TO DISPLAY AVAILABLE CATEGORIES IN UserInput FORM IN CatProp_ListView
                    //CurrCategories STORES CATEGORIES AS PropertyCategory (Navis API) TYPE
                    //THIS WILL BE ACCESSED IN STEP 2 (GetCatProperties()) AFTER USER HAS SELECTED WHICH CATEGORY TO ACCESS
                    CurrCategories.Add(oPC);
                    catDuplicate.Add(oPC.DisplayName);
                    ReturnCategories.Add(new Category
                    {
                        CatName = oPC.DisplayName
                    });
                }
                
                
            }
        }


        //STEP 2
        //ROUTED HERE FROM UserInput FORM AFTER USER SELECTS CATEGORY TYPE TO ACCESS ASSOCIATED PROPERTIES 
        //SELECTED CATEGORY IS PASSED AS CatNameSelected
        public static void GetCatProperties(string CatNameSelected)
        {
            //CYCLES THROUGH CATEGORY OF CLASS PropertyCategory
            //TO FIND MATCH THAT USER SELECTED OF TYPE STRING
            foreach (PropertyCategory category in CurrCategories)
            {
                if (category.DisplayName == CatNameSelected)
                {
                    //WHEN MATCH FOUND TAKES category OF CLASS PropertyCategory
                    //TO RETRIEVE AVAILABLE Properties
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


        //ROUTED FROM StarMain TO STORE PROJECT MODELS IN ModelList LIST 
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
    

    //CLASS BINDING CONTAINER FOR ReturnCategories Observable Collection
    public class Category
    {
        public string CatName { get; set; }
    }


    //CLASS BINDING CONTAINER FOR ReturnProp Observable Collection
    public class Property
    {
        public string PropName { get; set; }
        public string ValEx { get; set; }
    }
}
