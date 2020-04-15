using System;
using System.Collections.Generic;
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
using StartMain;

namespace SystemPropertyExporter
{
    /// <summary>
    /// Interaction logic for UserInput.xaml
    /// </summary>
    public partial class UserInput : Window
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

        public static int Selected_HierLvl { get; set; }

        public static string Selected_Cat { get; set; }

        public static string Selected_Prop { get; set; }


        //----------------------------------------------------------------------------------------


        public UserInput(string[] parameters)
        {
            InitializeComponent();

            //Makes models loaded in project (Selection Tree) visible to User in Models List View
            try
            {
                if (Start.firstOpen == true)
                {
                    Models_ComBox.ItemsSource = GetProperties.modelList;
                    Start.firstOpen = false;
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
                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
            }
            else if (CatRB.IsChecked == true && Models_ComBox.SelectedItem != null)
            {
                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
            }
            else if (ComponentRB.IsChecked == true && Models_ComBox.SelectedItem != null)
            {
                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
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
                    Selected_HierLvl = 1;
                    GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                    CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
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
                    Selected_HierLvl = 2;
                    GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                    CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
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
                    Selected_HierLvl = 3;
                    GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                    CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
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
                GetProperties.ReturnProp.Clear();

                //UPDATES AVAILABLE PROPERTIES WHEN CATEGORY SELECTED IN CatProp_ListView
                var selectedCat = CatProp_ListView.SelectedItem as GetProperties.Category;
                if (selectedCat != null)  //INITIATES PROPERTIES RETRIEVEL WHEN CATEGORY SELECTED (CONTAINER NOT EMPTY)
                {
                    GetProperties.GetCatProperties(selectedCat.CatName);
                    Selected_Cat = selectedCat.CatName;
                    Prop_ListView.ItemsSource = GetProperties.ReturnProp;
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
            int ui1 = -1;
            int ui2 = -2;
            int ui3 = -3;
            int ui4 = -4;
            bool duplicate = false;

            //Need to prevent item from being added if duplicate
            if (Dis_TB.Text == "INPUT DISCIPLINE MODEL" || Models_ComBox.Text == "SELECT MODEL"
                    || Selected_Cat == null || Selected_HierLvl < 1 || Selected_HierLvl > 3)
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
                        ui3 = combo.HierLvl.IndexOf(Selected_HierLvl.ToString());
                        ui4 = combo.SelectCat.IndexOf(Selected_Cat);

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
                    UserItems.Add(new Selected
                    {
                        Discipline = Dis_TB.Text,
                        ModFile = Models_ComBox.Text,
                        HierLvl = Selected_HierLvl.ToString(),
                        SelectCat = Selected_Cat
                    });

                    //DISPLAY USER SELECTION IN MODELSSELECTED_LISTVIEW
                    ModelsSelected_ListView.ItemsSource = UserItems;

                    //RESET SLECTION AND USER INPUTS
                    GetProperties.ReturnCategories.Clear();
                    GetProperties.ReturnProp.Clear();
                    Selected_Cat = null;
                    Dis_TB.Text = "INPUT DISCIPLINE MODEL";
                    Models_ComBox.Text = "SELECT MODEL";
                    Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
                } 
            }
        }
        

        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {   
            GetProperties.ReturnCategories.Clear();
            GetProperties.ReturnProp.Clear();
            Dis_TB.Text = "INPUT DISCIPLINE MODEL";
            Models_ComBox.Text = "SELECT MODEL";
            Dis_TB.Foreground = new SolidColorBrush(Color.FromRgb(169, 169, 169));
        }
        

        private void RemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Observable Collection (UserItems) bound to ModelSelected_Listview
                //Editable List 
                var selected = ModelsSelected_ListView.SelectedItems.Cast<object>().ToList();
                foreach(Selected item in selected)
                {
                    UserItems.Remove(item);
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
                GetProperties.modelList.Clear();
                this.Close();
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
}
