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
using SystemPropertyExporter;
using ClashData;

namespace StartMain
{
    //LOAD ADD-IN BUTTON IN NAVISWORKS RIBBON
    [PluginAttribute("StartMain.Start", //Namespace.Starting class of the plugin (where the override function is)
       "VDC.CAC",  // Your dev ID (It can be anything up to 7 letters I believe)
       ToolTip = "Model Data Export",    //Plugin Tooltip content
       DisplayName = "VDC Add-Ins")]    //Name of the plugin button.
    [RibbonLayout("AddinRibbon.xaml")]
    [RibbonTab("VDC Add-Ins")]
    [Command("Clash_Data_Exporter", Icon = "Data-Export-16.png", LargeIcon = "Data-Export-32.png", ToolTip = "Export Clash Detective Data to Excel for Power BI")]
    //SYSTEM PROPERTY EXPORTER - REFER TO AddinRibbon.xml FOR IMPLEMENTATION IN WPF
    [Command("System_Property_Exporter", Icon = "Prop-Export-16.ico", LargeIcon = "Prop-Export-32.ico", ToolTip = "Export MEP System Model Properties to Excel")]
    
    public class Start : CommandHandlerPlugin
        {
            public static bool FirstOpen { get; set; }

            public override int ExecuteCommand(string name, params string[] parameters)
            {
                try
                {
                    //BASED ON USER SELECTION, WILL REDIRECT TO DESIRED APPLICAITON
                    switch (name)
                    {
                        case "Clash_Data_Exporter":
                            Form1 form = new Form1(parameters);
                            form.ShowDialog();

                            form.Close();
                            break;

                        case "System_Property_Exporter":

                            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
                            GetPropertiesModel.DocModel = document.Models;

                            //IF FIRST TIME OPENING APP IN NAVISWORKS SESSION, PROMPTS UI TO LOAD DISCIPLINE MODELS IN APP
                            FirstOpen = true;

                            //IN THE EVENT NO MODELS EXIST IN PROJECT, APP TO PROMPT USER TO APPEND MODELS FIRST.
                            if (GetPropertiesModel.DocModel.Count == 0)
                            {
                                MessageBox.Show("No models currently appended in project." + "\n" + "Load models first.");
                                System.Windows.Application.Current.Shutdown();
                                break;
                            }
                            else
                            {
                                //RETRIEVES ALL BUILDING SYSTEM (DISICIPLINE) MODELS IN CURRENT PROJECT
                                //ASSIGNS VALUES (MODELS) TO ObservableCollection GetPropertiesModel.ModelList.
                                GetPropertiesModel.GetCurrModels();

                                //OPENS UserInput.xaml WINDOW.
                                UserInput ui = new UserInput(parameters);
                                ui.ShowDialog();
                            }
                            //ui.Close();
                            break;
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Error! Original Message: " + exception.Message);
                }

            return 0;
            }
    }
}
