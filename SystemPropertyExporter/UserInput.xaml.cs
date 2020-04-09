﻿using System;
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
            try
            {
                    Models_ComBox.ItemsSource = GetProperties.modelList;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }

        }

        private void SystemRB_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                //Reset Columns in Property Category and Property List View
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
                GetProperties.GetSystemProperties(Models_ComBox.SelectedItem.ToString(), "Layer");
                CatProp_ListView.ItemsSource = GetProperties.ReturnCategories;
                //hierarchyChange = true;
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
                //check previous selection is not same as current selection
                GetProperties.ReturnProp.Clear();

                //UPDATES AVAILABLE PROPERTIES WHEN CATEGORY SELECTED IN CatProp_ListView
                var selectedCat = CatProp_ListView.SelectedItem as GetProperties.Category;
                    if (selectedCat != null)  //INITIATES PROPERTIES RETRIEVEL WHEN CATEGORY SELECTED (CONTAINER NOT EMPTY)
                    {
                        GetProperties.GetCatProperties(selectedCat.CatName);
                        Prop_ListView.ItemsSource = GetProperties.ReturnProp;
                    }   
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error! Original Message: " + exception.Message);
            }

        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
        }
        
        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {   
            //CatProp_ListView.Items.Clear();
            GetProperties.ReturnCategories.Clear();
            GetProperties.ReturnProp.Clear();
        }
        
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}