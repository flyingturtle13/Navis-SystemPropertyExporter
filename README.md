# Navisworks Building System Properties Exporter
The ultimate goal of these tools and workflow is to improve MEP BIM coordination communication and federated model access with project teams.  The Navisworks add-in allows the user to select which discipline/system properties (per hierarchy level) is exported to an Excel spreadsheet for data interoperability.  The properties available to export are the same available in the "Properties" window.  The primary purpose is to send the system properties to Forge model viewer to create a 3D model metrics dashbaord to permit users to access system data, quantities, and statuses.  This workflow requires Autodesk Navisworks (Manage, Simulate, Freedom) 2019, Navisworks API to access model data for exporting and Forge model viewer.  However, this respository focuses on the Navisworks API and add-in only. 
</br></br>
**Note:** This add-in application is part of the VDC Add-Ins suite of tools and currently includes Clash Data Exporter (only works with Navisworks Manage).  For more information about the Clash Data Exporter tool see [Navisworks-Clash Data Exporter](https://github.com/flyingturtle13/Navis-Clash_Data_Exporter.git) here.

## Getting Started
Environment setup regarding application development logistics.

* IDE:
  * Visual Studio 2019
  
* Framework:
  * .NET Framework 4.7.2

* Language:
  * C#

* Output Type:
  * Dynamic-link Library (DLL)

* Additional Library Packages Implemented: </br>
  * Navisworks API </br>
    - Autodesk.Navisworks.Api
    - Autodesk.Navisworks.Automation
    - Autodesk.Navisworks.Clash
    - navisworks.gui.roamer
    - AdWindows
  * Microsoft.Office.Interop.Excel

## Application Development
Application features and specs for Navisworks Manage add-in

* Software Required
  - Navisworks 2019 (Manage, Simulate, Freedom)
    * If other version desired, replace Autodesk library packages with Navisworks version desired in Visual Studio. 
    
* Navisworks Model Building System Properties Accessed
  - Parameters being exported can be seen by viewing the Selection Tree window and Properties window.
  - Hierarchy Levels 
    * Discipline System - File icon
    * System Parts - Layer icon (exported from AutoCAD) or Collection icon (exported from Revit)
    * Individual Components - Block icon or Geometry icon (directly sub item to a Layer)
  - Navisworks Selection Tree - Models Exported from Revit
    <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80930464-99858100-8d68-11ea-9de0-32a8c4be8fb6.png" width="250">
    </p>
  - Navisworks Selection Tree - Models Exported from AutoCAD
    <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80930540-0e58bb00-8d69-11ea-96e8-ad4ee28e3b59.png" width="250">
    </p>
  - Navisworks Properties Window
    <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80930589-6099dc00-8d69-11ea-836d-4c50385cb416.png" width="300">
    </p>

* User Interface
  - User to input Discipline / Building System
  - User to select associated model (similar to what is seen in the Selection Tree)
  - User to select model hierarchy level to export (entire system, system parts, individual components)
  - User to select which Property Category to export parameters
  - Creates a viewable list of property categories per discipline to be exported
  - User has ability to save and load a list of discipline property categories to export to eliminate list creation every time export is needed
  - Exported parameters are stored in an Excel file that the user can choose save location

## Application Structure
Overall add-in process flowchart.  User interface in Navisworks is included for reference.
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/80936145-aca64a00-8d84-11ea-84ee-448722c0288b.png" width="1000">
</p>
       
## Navisworks API Implementation
Below highlights specific API features implemented to access and export specific Clash Detective Data
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/80936415-f5aace00-8d85-11ea-93a1-32b978208db2.png" width="700">
</p>

##### Hierarchy Level Mapping Visual Using Selection Tree
- How API is mapped to model files in Selection Tree UI
  - Model files exported from Revit
    <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80939455-88049f00-8d91-11ea-997d-afbf994902b7.png" width="700">
    </p>
  - Model files exported from AutoCAD
    <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80939489-a79bc780-8d91-11ea-8a75-d36d85c60939.png" width="700">
    </p>

##### Model Categories and Properties Mapping Visual Using Properties Window
- How API is mapped to model files to retrieve available categories and properties for export
  <p align="center">
     <img src="https://user-images.githubusercontent.com/44215479/80939719-5cce7f80-8d92-11ea-801b-4a4c7946e9fa.png" width="650">
  </p>

##### Output Excel Spreadsheet Example
- Export Excel file example based on user input mapping API classes and properties
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/80940869-c603c200-8d95-11ea-9afa-f5bf6e7fdc97.png" width="1000">
</p>

## Installing and Running Add-in
1. Clone or download project. </br>
2. Open SystemPropertyExporter.sln in Visual Studio 2019. </br>
   * **Note:** SystemPropertyExporter.zip in Add-In Example Files contains example of working add-in for Navisworks 2019. Unzip and copy into **Local_Drive:\...\Autodesk\Navisworks Manage 2019\Plugins** if one would like to test without opening Visual Studio
3. Ensure that the library packages stated in Getting Started are installed and referenced. </br>
4. The application can then be run in debug mode. </br>
5. Go to Debug/Release location and copy the files and folder below. </br>
   <p align = "center">
      <img src="https://user-images.githubusercontent.com/44215479/80941253-d6686c80-8d96-11ea-9c51-177ce4ddb2a9.png">
   </p>
6. Create a SystemPropertyExporter folder in **Local_Drive:\...\Autodesk\Navisworks Manage 2019\Plugins** </br>
7. Paste copied files and folders to SystemPropertyExporter folder </br>
   * **Note:** This application is part of the VDC Add-Ins suite, so include ClashData.dll in SystemPropertyExporter folder (included SystemPropertyExporterl.zip in Add-In Example Files folder)
   * **Note:** If user would like to remove Clash Data Add-In, can be modified in StartMain.cs prior to debugging and release
8. Open Navisworks Manage 2019 (or whichever version if Autodesk references modified) to execute and test add-in.

## Add-In Implementation User Instructions
1) Select System Property Exporter (In reibbon: VDC Add-Ins --> Export Tools --> System Property Exporter) 
   <p align = "center">
      <img src="https://user-images.githubusercontent.com/44215479/80942187-06b10a80-8d99-11ea-903e-dae92a6e841a.png" width="600">
   </p>
2) User to input or make selections to export desired system properties. See below for a discription of each feature is described below.
   <p align = "center">
      <img src="https://user-images.githubusercontent.com/44215479/80942605-0bc28980-8d9a-11ea-8c1c-834d1598e54d.png" width="800">
   </p></br>
 * (1) User to input name of building system and discipline organization code.
 * (2) Select associated model file from drop-down list.
 * (3) Select at which system level of information to export. Typically, System Individual Components is selected to get the most detail of each system element.
 * (4) User to select which Category in the list properties and values will be exported.
 * (5) An preview of what kind of available properties and values are displayed.
 * (6) Add button will include in queue of properties to be exported.
 * (7) Reset button will clear user modifications.
 * (8) The list previews all user items to be exported (based on user selecting the Add button) seen in Selected Models & Properties to Export.
 * (9) Remove button will remove any items seleted in the list in (8). </br>
 * (10) Save List button allows the user to save the list (.TXT format) that has been created in (8) for future automated population. 
 * (11) Load List button allows the user to load a previously saved list.
 * (12) OK button executes items in Selected Models & Properties Export list to be recorded to an Excel spreadsheet.
   - User will be prompted where to save when complete
   - Application will automatically close when complete
 * (13) Cancel button will close application without exporting any system properties.
   
3) Working UI example
   <p align = "center">
      <img src="https://user-images.githubusercontent.com/44215479/80944803-36631100-8d9f-11ea-814c-b950e1e2df97.png" width="800">
   </p>
   
## References for Further Learning
- [Customizing Autodesk® Navisworks® 2013 with the .NET API - Simon Bee](https://www.autodesk.com/autodesk-university/class/Customizing-AutodeskR-NavisworksR-2013-NET-API-2012)
- [Navisworks .NET API Properties - Xiaodong Liang](https://adndevblog.typepad.com/aec/2012/05/navisworks-net-api-properties.html)
- [API Docs - Guilherme Talarico](https://apidocs.co/apps/navisworks/2018/87317537-2911-4c08-b492-6496c82b3ed1.htm#)

