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

namespace SystemPropertyExporter
{
    /// <summary>
    /// Interaction logic for UserInput.xaml
    /// </summary>
    public partial class UserInput : Window
    {
        public UserInput(string[] parameters)
        {
            InitializeComponent();

            //Makes models loaded in project (Selection Tree) visible to User in Models List View
            //Models_ListView.SelectedItem = true;
            try
            {
                    Models_ComBox.ItemsSource = GetProperties.modelList;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }

        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            if (SystemRB.IsChecked == true)
            {

            }
            else if (CatRB.IsChecked == true)
            {

            }
            else if (ComponentRB.IsChecked == true)
            {

            }
        }

        private void SystemRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                //Reset Columns in Property Category and Property List View
                //if (CatProp_ListView.SelectedItem != null)
                //{
                //    CatProp_ListView.SelectedItem = null;
                //    CatProp_ListView.SelectedIndex = 0;
                //}
                CatProp_ListView.ItemsSource = null;

                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "File");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }
        
        private void CatRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                //Reset Columns in Property Category and Property List View
                //if (CatProp_ListView.SelectedItem != null)
                //{
                //    CatProp_ListView.SelectedItem = null;
                //    CatProp_ListView.SelectedIndex = 0;
                //}
                CatProp_ListView.ItemsSource = null;
                
                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
}
        
        private void ComponentRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                //Reset Columns in Property Category & Property List View
                //if (CatProp_ListView.SelectedItem != null)
                //{
                //    CatProp_ListView.SelectedItem = null;
                //    CatProp_ListView.SelectedIndex = 0;
                //}
                CatProp_ListView.ItemsSource = null;

                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Block");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }
        }

        private void PropCat_Selection(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                
                GetProperties.ReturnProp.Clear();
                Prop_ListView.ItemsSource = null;

                var selectedItem = CatProp_ListView.SelectedItem as GetProperties.Category;
                GetProperties.GetCatProperties(selectedItem.CatName);
                Prop_ListView.ItemsSource = GetProperties.ReturnProp;
                //if (SystemRB.IsChecked == true || CatRB.IsChecked == true || ComponentRB.IsChecked == true)
                //{
                //    CatProp_ListView.SelectedItem = null;
                //}
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }

        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
