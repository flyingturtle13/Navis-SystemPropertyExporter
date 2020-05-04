using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SystemPropertyExporter;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using StartMain;

namespace SystemPropertyExporter
{
    /// <summary>
    /// Interaction logic for UserInput.xaml
    /// </summary>
    public partial class UserInput : Window
    {
        public UserInput(string[]parameters)
        {
            InitializeComponent();

            //MAKES MODELS LOADED IN PROJECT (NAVISWORKS SELECTION TREE) VISIBLE TO USER IN PROJECT MODELS ComboBox
            if (Start.FirstOpen == true)
            {
                Models_ComBox.ItemsSource = GetPropertiesModel.ModelList;
                Start.FirstOpen = false;
            }
        }


        //----------------------------------------------------------------------------------------------------------


        //MODEL FILES COMBO BOX - USER TO INITIALLY SELECT SPECIFIC HIERARCHY LEVEL//
        //WHEN NEW SELECTION MADE IN RADIO BUTTON GROUP, RETRIEVES AVAILABLE CATEGORIES (GetProertiesModel.cs)
        //& RELOADS IN CatProp_LISTVIEW
        //USES OBSERVABLE COLLECTION PROPERTIES.
        private void ModelCB_Select(object sender, SelectionChangedEventArgs e)
        {
            if (SystemRB.IsChecked == true && Models_ComBox.SelectedItem != null)
            {
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
            else if (CatRB.IsChecked == true && Models_ComBox.SelectedItem != null)
            {
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
            else if (ComponentRB.IsChecked == true && Models_ComBox.SelectedItem != null)
            {
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
        }


        //HIERARCHY LEVEL SELECTION RADIO BUTTON - USER TO INITIALLY SELECT//
        //GROUP NAME = HIERARCHY//
        private void SystemRB_Checked(object sender, RoutedEventArgs e)
        {
            //MUST CHECK IF USER HAS SELECTED A MODEL FROM Models_ComBox TO
            //DETERMINE IF CATEGORIES SHOULD BE RETRIEVED FROM GetPropertiesModel.GetSystemProperties
            if (Models_ComBox.SelectedItem != null)
            {
                ExportProperties.Selected_HierLvl = 1;
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
        }


        //HIERARCHY LEVEL SELECTION RADIO BUTTON - USER TO INITIALLY SELECT//
        //GROUP NAME = HIERARCHY//
        private void CatRB_Checked(object sender, RoutedEventArgs e)
        {
            //MUST CHECK IF USER HAS SELECTED A MODEL FROM Models_ComBox TO
            //DETERMINE IF CATEGORIES SHOULD BE RETRIEVED FROM GetPropertiesModel.GetSystemProperties
            if (Models_ComBox.SelectedItem != null)
            {
                ExportProperties.Selected_HierLvl = 2;
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
        }


        //HIERARCHY LEVEL SELECTION - USER TO INITIALLY SELECT//
        //GROUP NAME = HIERARCHY//
        private void ComponentRB_Checked(object sender, RoutedEventArgs e)
        {
            //MUST CHECK IF USER HAS SELECTED A MODEL FROM Models_ComBox TO
            //DETERMINE IF CATEGORIES SHOULD BE RETRIEVED FROM GetPropertiesModel.GetSystemProperties
            if (Models_ComBox.SelectedItem != null)
            {
                ExportProperties.Selected_HierLvl = 3;
                GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
            }
        }


        //DISCIPLINE NAME TEXT BOX - USER TO INITIALLY INPUT//
        private void Dis_KeyDn(object sender, KeyEventArgs e)
        {
            //IF DEFAULT TEXT IN TEXT BOX, WILL CLEAR DEFAULT TEXT
            //TO ACCOMODATE USER INPUT WHEN SELECTED
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL")
            {
                Dis_TB.Text = "";
                Dis_TB.Foreground = Brushes.Black;
            }
        }


        //DISCIPLINE NAME TEXT BOX - USER TO INITIALLY INPUT//
        private void Dis_GotFocus(object sender, RoutedEventArgs e)
        {
            //IF DEFAULT TEXT IN TEXT BOX, WILL CLEAR DEFAULT TEXT
            //TO ACCOMODATE USER INPUT WHEN SELECTED
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL")
            {
                Dis_TB.Text = "";
                Dis_TB.Foreground = Brushes.Black;
            }
        }


        //IF USER DESELECTS TEXT BOX, INITIATES METHOD
        private void Dis_LostFocus(object sender, RoutedEventArgs e)
        {
            //IF USER LEAVES TEXT BOX BLANK, WILL RE-DISPLAY DEFAULT INSTRUCTION TEXT
            if (Dis_TB.Text == "")
            {
                Dis_TB.Text = "INPUT DISCIPLINE MODEL";
                Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
            }
        }


        //CATEGORY LISTVIEW - USER TO INITIALLY SELECT//
        private void PropCat_Selection(object sender, SelectionChangedEventArgs e)
        {
            //CHECK PREVIOUS SELECTION IS NOT SAME AS CURRENT SELECTION
            GetPropertiesModel.ReturnProp.Clear();

            //UPDATES AVAILABLE PROPERTIES WHEN CATEGORY SELECTED IN CatProp_ListView
            var selectedCat = CatProp_ListView.SelectedItem as Category;

            if (selectedCat != null)  //INITIATES PROPERTIES RETRIEVEL WHEN CATEGORY SELECTED (CONTAINER NOT EMPTY)
                                      //OTHERWISE UPDATES CATEGORIES SINCE NO SelectedItem IS BOUND
            {
                //RESPECTS OBSERVABLE COLLECTION PROPERTIES
                GetPropertiesModel.GetCatProperties(selectedCat.CatName); //TAKES USER SELECTED CATEGORY TO RETRIEVE PROPERTIES
                ExportProperties.Selected_Cat = selectedCat.CatName;
                Prop_ListView.ItemsSource = GetPropertiesModel.ReturnProp;
            }
        }


        //---------------------------------------------------------------------------------------------------------------------


        //ADD BUTTON - QUEUES USER MODEL PROPERTIES FOR EXPORT
        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            int ui1;
            int ui2;
            int ui3;
            int ui4;
            bool duplicate = false;

            //CHECKS THAT MODELSELECTED_LISTVIEW IS NOT EMPTY
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL" || Models_ComBox.Text == "SELECT MODEL"
                    || ExportProperties.Selected_Cat == null || ExportProperties.Selected_HierLvl < 1 
                    || ExportProperties.Selected_HierLvl > 3)
            {
                MessageBox.Show("All fields require input or selection.");
            }
            else
            {
                //PREVENTS ITEM FROM BEING ADDED IF ALREADY EXISTS IN LIST(PREVENTS DUPLICATES FROM BEING ADDED TO LIST)
                
                //FIRST, CONFIRMS EXPORT LIST IS NOT NULL
                if (ModelsSelected_ListView != null)
                {
                    //OBSERVABLE COLLECTION (UserItems) BOUND TO ModelSelected_Listview
                    //CREATES EDITABLE LIST THAT CAN BE MANIPULATED
                    var cUserItems = ModelsSelected_ListView.Items.Cast<object>().ToList();

                    //CHECKS IF NEW INPUT WILL BE A DUPLICATE
                    foreach (Selected combo in cUserItems)
                    {
                        ui1 = combo.Discipline.IndexOf(Dis_TB.Text);
                        ui2 = combo.ModFile.IndexOf(Models_ComBox.Text);
                        ui3 = combo.HierLvl.IndexOf(ExportProperties.Selected_HierLvl.ToString());
                        ui4 = combo.SelectCat.IndexOf(ExportProperties.Selected_Cat);

                        if (ui1 == ui2 && ui2 == ui3 && ui3 == ui4)
                        {
                            MessageBox.Show("This export combination already exists. \n Duplicate will not be added.");
                            duplicate = true;
                            break;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Must add items to list to export properties.");
                }

                //IF PASSES DUPLICATES TEST, USER INPUTS ARE ADDED TO TEMPORARY ITEMS EXPORT QUEUE LIST CONTAINER (UserItems)
                if (duplicate == false)
                {
                    ExportProperties.UserItems.Add(new Selected
                    {
                        Discipline = Dis_TB.Text,
                        ModFile = Models_ComBox.Text,
                        HierLvl = ExportProperties.Selected_HierLvl.ToString(),
                        SelectCat = ExportProperties.Selected_Cat
                    });

                    //DISPLAYS UPDATED USER SELECTION FOR EXPORT IN MODELSSELECTED_LISTVIEW
                    ModelsSelected_ListView.ItemsSource = ExportProperties.UserItems;

                    //RESETS SLECTION AND USER INPUTS
                    GetPropertiesModel.ReturnCategories.Clear();
                    GetPropertiesModel.ReturnProp.Clear();
                    ExportProperties.Selected_Cat = null;
                    Dis_TB.Text = "INPUT DISCIPLINE MODEL";
                    Models_ComBox.Text = "SELECT MODEL";
                    Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
                } 
            }
        }
        

        //RESET BUTTON - CLEARS ANY USER INPUT MANIPULATION TO DEFAULT WHEN APP ORIGINALLY OPEN//
        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {   
            GetPropertiesModel.ReturnCategories.Clear();
            GetPropertiesModel.ReturnProp.Clear();
            Dis_TB.Text = "INPUT DISCIPLINE MODEL";
            Models_ComBox.Text = "SELECT MODEL";
            Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
        }


        //SAVE BUTTON - ABILITY TO SAVE USER LIST SO DOES NOT HAVE TO BE
        //RECREATED FOR FUTURE RE-EXPORTING ACTIVITIES.
        private void SaveList_Click(object sender, RoutedEventArgs e)
        {
            //CREATES NEW VARIABLE INSTANCE - ALLOWS TO OPEN WINDOWS EXPLORER SAVE FILE PROMPT
            string filename = "";
            SaveFileDialog saveList = new SaveFileDialog();

            //SETS FILE TYPE TO BE SAVED AS .TXT
            saveList.Title = "Save to...";
            saveList.Filter = "Text Documents | *.txt";

            //OPENS WINDOWS EXPLORER TO BEGIN LIST SAVE PROCESS
            if (saveList.ShowDialog() == true)
            {
                filename = saveList.FileName.ToString();

                //CHECKS USER HAS INPUTED A NAME FOR THE FILE
                if (filename != "")
                {
                    using (StreamWriter sw = new StreamWriter(filename))
                    {
                        //POPULATES TXT FILE WITH LIST ITEMS SEPARATED BY "--" USING StreamWriter & CLOSES WHEN COMPLETE
                        var selected = ModelsSelected_ListView.ItemsSource.Cast<object>().ToList();
                        foreach (Selected item in selected)
                        {
                            sw.WriteLine("--");
                            sw.WriteLine(item.Discipline);
                            sw.WriteLine(item.ModFile);
                            sw.WriteLine(item.HierLvl);
                            sw.WriteLine(item.SelectCat);
                        }
                        sw.Dispose();
                        sw.Close();
                    }
                }
            }
        }
        

        //LOAD LIST BUTTON - ABILITY TO LOAD A PREVIOUSLY SAVED LIST
        //SO USER LIST DOES NOT HAVE TO BE RECREATED MANUALLY.
        private void LoadList_Click(object sender, RoutedEventArgs e)
        {
            //CREATES NEW VARIABLE INSTANCE - ALLOWS TO OPEN WINDOWS EXPLORER OPEN A FILE
            string filename = "";
            OpenFileDialog loadList = new OpenFileDialog();

            //SETS FILE TYPE THAT CAN BE OPENED (.TXT)
            loadList.Title = "Open File";
            loadList.Filter = "Text Documents | *.txt";
            
            //OPENS WINDOWS EXPLORER TO BEGIN LIST OPEN FILE PROCESS
            if (loadList.ShowDialog() == true)
            {
                try
                {
                    //CLEARS ANY PREVIOUS ITEMS ADDED TO TEMPORARY LIST CONTAINER FOR EXPORT
                    ExportProperties.UserItems.Clear();

                    //BEGINS READING .TXT FILE 
                    filename = loadList.FileName.ToString();
                    var fileLines = File.ReadAllLines(filename);

                    int i = 0;

                    //READS TXT FILE LINE BY LINE 
                    foreach (String line in fileLines)
                    {
                        //IF "--" IS READ, WILL ADD A NEW ROW OF ASSOCIATED INPUTS TO UserItems TEMPORARY CONTAINER LIST TO BE DISPLAYED
                        if (line == "--")
                        {
                            ExportProperties.UserItems.Add( new Selected
                            {
                                Discipline = fileLines[i+1],
                                ModFile = fileLines[i+2],
                                HierLvl = fileLines[i+3],
                                SelectCat = fileLines[i+4]
                            });
                        }

                        i++;
                    }

                    //DISPLAYS ADDED ITEMS IN UI
                    ModelsSelected_ListView.ItemsSource = ExportProperties.UserItems;
                }
                catch (Exception x)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + x.Message);
                }
            }
        }


        //REMOVE BUTTON - DELETES ADDED ITEM TO ModelsSelected_ListView IF USER CHOOSES
        //USER MUST SELECT ONE OR MULTIPLE ITEMS TO BE REMOVED.
        private void RemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            //OBSERVABLE COLLECTION (UserItems) BOUND TO ModelSelected_ListView
            //CREATES A COPY SO LIST CAN BE ITERATED OVER AND ISOLATE/SELECT AN ITEM
            var selected = ModelsSelected_ListView.SelectedItems.Cast<object>().ToList();
            foreach(Selected item in selected)
            {
                ExportProperties.UserItems.Remove(item);
            }
        }


        //CANCEL BUTTON - CLOSES APPLICATION WITHOUT PERFORMING EXPORT PROCESS//
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            GetPropertiesModel.ModelList.Clear();
            ExportProperties.Idx = 0;
            this.Close();
        }


        //OK BUTTON - EXECUTES EXPORT PROCESS AND CLOSES APP AUTOMATICALLY WHEN COMPLETE//
        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            //DIRECT TO ExportProperties CLASS FOR SEARCHING AND STORING DESIRED
            //CATEGORY, PROPERTIES, AND VALUES.
            //INITIALIZES Idx TO 0 FOR ASSOCIATING PROPERTIES/VALUES WITH
            //ITEMS WITH ALL ITEMS FROM UserItems AND ExportItems.
            ExportProperties.Idx = 0;
            ExportProperties.ProcessModelsSelected();

            //RESET LISTS AND OBSERVABLE COLLECTIONS  BEFORE CLOSING PLUG-IN
            GetPropertiesModel.ModelList.Clear();
            ExportProperties.ExportItems.Clear();
            ExportProperties.UserItems.Clear();
            ExportProperties.ExportProp.Clear();
            ExportProperties.ExportVal.Clear();
            ExportProperties.ItemIdx.Clear();

            //TERMINATES ADD-IN AND CLOSES WINDOW
            this.Close();
        }

        
    }
}
