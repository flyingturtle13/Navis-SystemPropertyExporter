# Navisworks Building System Property Exporter
The ultimate goal of these tools and workflow is to improve MEP BIM coordination communication and federated model access with project teams.  The Navisworks add-in allows the user to select which discipline/system properties (per hierarchy level) is exported to an Excel spreadsheet for data interoperability.  The properties available to export are the same available in the "Properties" window.  The primary purpose is to send the system properties to Forge model viewer to create a 3D model metrics dashbaord to permit users to access system data, quantities, and statuses.  This workflow requires Autodesk Navisworks (Manage, Simulate, Freedom) 2019, Navisworks API to access model data for exporting and Forge model viewer.  However, this respository focuses on the Navisworks API and add-in only.

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
  <img src="https://user-images.githubusercontent.com/44215479/80930142-63470200-8d66-11ea-836a-64439e02f72d.png" width="1000">
</p>
       
## Navisworks API Implementation
Below highlights specific API features implemented to access and export specific Clash Detective Data
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75583929-7ca89680-5a23-11ea-99a4-8d49d07a47e4.png" width="600">
</p>

##### Total Objects By Discipline Module
- How API is mapped to model files in Selection Tree UI
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75584908-b8446000-5a25-11ea-99ed-18e675b9782d.png" width="600">
</p>

##### Output Excel Spreadsheet Examples
- Export from Clash Test module
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75585249-77991680-5a26-11ea-9e63-43ed83651cb1.png" width="1000">
</p>

- Export from Total Objects by Discipline module
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75585302-9ac3c600-5a26-11ea-950d-f41adfa4aa90.png" width="400">
</p>

## Installing and Running Application
1. Clone or download project. </br>
2. Open ClashData.sln in Visual Studio 2019. </br>
3. Ensure that the library packages stated in Getting Started are installed and referenced. </br>
4. The application can then be run in debug mode. </br>
5. Go to Debug/Release location and copy the files and folder below. </br>
   <p align = "center">
      <img src="https://user-images.githubusercontent.com/44215479/75586032-50434900-5a28-11ea-9cb2-8e3d00008d17.png" width = "200">
   </p>
6. Create a ClashData folder in **Local_Drive:\...\Autodesk\Navisworks Manage 2019\Plugins** </br>
7. Paste copied files and folders to ClashData folder </br>
8. Open Navisworks Manage 2019 to execute and test add-in.

## References for Further Learning
- Tools and Workflow described here are based on AU 2019 Presentation: [Visualizing Clash Metrics in Navisworks with Power BI - Carlo Caparas](https://www.autodesk.com/autodesk-university/class/Its-All-Data-Visualizing-Clash-Metrics-Navisworks-and-Power-BI-2019)
- [Customizing Autodesk® Navisworks® 2013 with the .NET API - Simon Bee](https://www.autodesk.com/autodesk-university/class/Customizing-AutodeskR-NavisworksR-2013-NET-API-2012)
- [Navisworks .NET API 2013 new feature – Clash 1 - Xiaodong Liang](https://adndevblog.typepad.com/aec/2012/05/navisworks-net-api-2013-new-feature-clash-1.html)
- [Navisworks .NET API 2013 new feature – Clash 2 - Xiaodong Liang](https://adndevblog.typepad.com/aec/2012/05/navisworks-net-api-2013-new-feature-clash-2.html)
- [API Docs - Guilherme Talarico](https://apidocs.co/apps/navisworks/2018/87317537-2911-4c08-b492-6496c82b3ed1.htm#)
- [Power BI Documentation - Microsoft Corporation](https://docs.microsoft.com/en-us/power-bi/#pivot=home&panel=home-all)

