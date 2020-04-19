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

            //Makes models loaded in project (Selection Tree) visible to User in Models List View
            try
            {
                if (Start.FirstOpen == true)
                {
                    Models_ComBox.ItemsSource = GetPropertiesModel.ModelList;
                    Start.FirstOpen = false;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        //--------------------------------------------------------------------------


        //MODEL FILES COMBO BOX - USER TO INITIALLY SELECT//
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
            try
            {
                if (Models_ComBox.SelectedItem != null)
                {
                    ExportProperties.Selected_HierLvl = 1;
                    GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                    CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
                }

            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        //HIERARCHY LEVEL SELECTION RADIO BUTTON - USER TO INITIALLY SELECT//
        //GROUP NAME = HIERARCHY//
        private void CatRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Models_ComBox.SelectedItem != null)
                {
                    ExportProperties.Selected_HierLvl = 2;
                    GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                    CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        //HIERARCHY LEVEL SELECTION - USER TO INITIALLY SELECT//
        //GROUP NAME = HIERARCHY//
        private void ComponentRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Models_ComBox.SelectedItem != null)
                {
                    ExportProperties.Selected_HierLvl = 3;
                    GetPropertiesModel.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                    CatProp_ListView.ItemsSource = GetPropertiesModel.ReturnCategories;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        //DISCIPLINE NAME TEXT BOX - USER TO INITIALLY SELECT//
        private void Dis_KeyDn(object sender, KeyEventArgs e)
        {
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL")
            {
                Dis_TB.Text = "";
                Dis_TB.Foreground = Brushes.Black;
            }
        }


        private void Dis_GotFocus(object sender, RoutedEventArgs e)
        {
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL")
            {
                Dis_TB.Text = "";
                Dis_TB.Foreground = Brushes.Black;
            }
        }


        private void Dis_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Dis_TB.Text == "")
            {
                Dis_TB.Text = "INPUT DISCIPLINE MODEL";
                Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
            }
        }


        //CATEGORY LISTVIEW - USER TO INITIALLY SELECT//
        private void PropCat_Selection(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                //check previous selection is not same as current selection
                GetPropertiesModel.ReturnProp.Clear();

                //UPDATES AVAILABLE PROPERTIES WHEN CATEGORY SELECTED IN CatProp_ListView
                var selectedCat = CatProp_ListView.SelectedItem as Category;
                if (selectedCat != null)  //INITIATES PROPERTIES RETRIEVEL WHEN CATEGORY SELECTED (CONTAINER NOT EMPTY)
                {
                    GetPropertiesModel.GetCatProperties(selectedCat.CatName);
                    ExportProperties.Selected_Cat = selectedCat.CatName;
                    Prop_ListView.ItemsSource = GetPropertiesModel.ReturnProp;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }

        }


        //-------------------------------------------------------------------------------


        //ADD BUTTON - QUEUES USER MODEL PROPERTIES FOR EXPORT
        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            int ui1;
            int ui2;
            int ui3;
            int ui4;
            bool duplicate = false;

            //Need to prevent item from being added if duplicate
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL" || Models_ComBox.Text == "SELECT MODEL"
                    || ExportProperties.Selected_Cat == null || ExportProperties.Selected_HierLvl < 1 
                    || ExportProperties.Selected_HierLvl > 3)
            {
                MessageBox.Show("All fields require input or selection.");
            }
            else
            {
                //CHECKS THAT MODELSELECTED_LISTVIEW IS NOT EMPTY
                if (ModelsSelected_ListView != null)
                {
                    //OBSERVABLE COLLECTION (UserItems) BOUND TO ModelSelected_Listview
                    //EDITABLE LIST
                    var cUserItems = ModelsSelected_ListView.Items.Cast<object>().ToList();
                    foreach (Selected combo in cUserItems)
                    {
                        ui1 = combo.Discipline.IndexOf(Dis_TB.Text);
                        ui2 = combo.ModFile.IndexOf(Models_ComBox.Text);
                        ui3 = combo.HierLvl.IndexOf(ExportProperties.Selected_HierLvl.ToString());
                        ui4 = combo.SelectCat.IndexOf(ExportProperties.Selected_Cat);

                            //MessageBox.Show($"{combo.Discipline},{combo.ModFile},{combo.HierLvl},{combo.SelectCat}");
                            //MessageBox.Show($"{ui1},{ui2},{ui3},{ui4}");

                        if (ui1 == ui2 && ui2 == ui3 && ui3 == ui4)
                        {
                            MessageBox.Show("This export combination already exists. \n Duplicate will not be added.");
                            duplicate = true;
                            break;
                        }
                    }
                } 

                if (duplicate == false)
                {
                    ExportProperties.UserItems.Add(new Selected
                    {
                        Discipline = Dis_TB.Text,
                        ModFile = Models_ComBox.Text,
                        HierLvl = ExportProperties.Selected_HierLvl.ToString(),
                        SelectCat = ExportProperties.Selected_Cat
                    });

                    //DISPLAY USER SELECTION IN MODELSSELECTED_LISTVIEW
                    ModelsSelected_ListView.ItemsSource = ExportProperties.UserItems;

                    //RESET SLECTION AND USER INPUTS
                    GetPropertiesModel.ReturnCategories.Clear();
                    GetPropertiesModel.ReturnProp.Clear();
                    ExportProperties.Selected_Cat = null;
                    Dis_TB.Text = "INPUT DISCIPLINE MODEL";
                    Models_ComBox.Text = "SELECT MODEL";
                    Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
                } 
            }
        }
        

        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {   
            GetPropertiesModel.ReturnCategories.Clear();
            GetPropertiesModel.ReturnProp.Clear();
            Dis_TB.Text = "INPUT DISCIPLINE MODEL";
            Models_ComBox.Text = "SELECT MODEL";
            Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
        }


        //SAVE BUTTON - ABILITY TO SAVE USER LIST SO DOES NOT HAVE TO BE
        //RECREATED FOR FUTURE USE.
        private void SaveList_Click(object sender, RoutedEventArgs e)
        {
            string filename = "";
            SaveFileDialog saveList = new SaveFileDialog();

            saveList.Title = "Save to...";
            saveList.Filter = "Text Documents | *.txt";

            if (saveList.ShowDialog() == true)
            {
                filename = saveList.FileName.ToString();

                if (filename != "")
                {
                    using (StreamWriter sw = new StreamWriter(filename))
                    {
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
        //SO USER LIST DOES NOT HAVE TO BE RECREATED.
        private void LoadList_Click(object sender, RoutedEventArgs e)
        {
            string filename = "";
            OpenFileDialog loadList = new OpenFileDialog();

            loadList.Title = "Open File";
            loadList.Filter = "Text Documents | *.txt";

            if (loadList.ShowDialog() == true)
            {
                try
                {
                    ExportProperties.UserItems.Clear();

                    filename = loadList.FileName.ToString();
                    var fileLines = File.ReadAllLines(filename);

                    int i = 0;

                    foreach (String line in fileLines)
                    {
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
            try
            {
                //Observable Collection (UserItems) bound to ModelSelected_Listview
                //Editable List 
                var selected = ModelsSelected_ListView.SelectedItems.Cast<object>().ToList();
                foreach(Selected item in selected)
                {
                    ExportProperties.UserItems.Remove(item);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GetPropertiesModel.ModelList.Clear();
                this.Close();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }


        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            ExportProperties.ProcessModelsSelected();
            this.Close();
        }

        
    }
}
