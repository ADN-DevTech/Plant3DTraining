
using System;
using System.Collections.Generic;


// **************** Start of LAB 1 **************************


// 1.1 Add References to the following .NET assemblies
// acdbmgd.dll (AutoCAD - Set Copy Local to False)
// acmgd.dll   (AutoCAD - Set Copy Local to False)
// accoremgd.dll   (AutoCAD - Set Copy Local to False)
// PnPProjectManagerMgd.dll (Plant 3D - P&ID)
// PnIdProjectPartsMgd.dll  

// for getLineGroup() - "LineGroup"
// Add a reference to PnIDMgd.dll (needed for LineSegment) 
// Add a reference to PnPCommonDbxMgd (needed for LineStyle)
// Add a reference to PnPDataLinks.dll (needed for DataLinksManager)

// Add a reference to PnP3dObjectsMgd.dll needed for ConnectionIterator //LAB 4 

// PnP3dProjectPartsMgd.dll  // LABS 1 PipingProject 
// PnP3dConnectionManager // LAB 4 ConnectionManager



// 1.2 Add using statements to bring in the following namespaces 
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance; // (needed to use PlantApplication and PlantProject)

// Add a using statement to bring in this namespace
using Autodesk.ProcessPower.P3dProjectParts; // PnP3dProjectPartsMgd.dll  // LAB 4 PipingProject


// 1.3 use the Using statement to create a variable AcadApp as 
// the Autodesk.AutoCAD.ApplicationServices.Application;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;


namespace LABS_Complete
{
    public class LAB_01
    {
        // Using the Project
        // The Project class contains services to access project level information. 
        // This includes the project settings, the DWG files in the project, and 
        // links to other data files and objects such as the symbol & style library, 
        // the data links manager to access the project database information, etc.
        // To use the class, add a reference to the PnPProjectManagerMgd.dll .NET assembly. 
        // NOTE: At this time there are no APIs to control Project Manager UI level operations.
        // Below is code that asks for the current project object and then lists 
        // all the drawings of the project in the AutoCAD console window.
        // C# Code


        // 1.4 Use the CommandMethod attribute and create a command "ProjectStructure"
        // use something like Project_Structure for the function name
        [CommandMethod("ProjectStructure")] 
        public static void Project_Structure()
        {
            // 1.5 Instantiate an variable as an Editor using
            // AcadApp.DocumentManager.MdiActiveDocument.Editor; 
            // (AcadApp is the name of the variable from step 3)
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // 1.6 Declare a variable as a PlantProject. Instantiate it using
            // the CurrentProject of PlantApplication. (PlantApplication is a singleton)  
            PlantProject currentProj = PlantApplication.CurrentProject;


            // 1.7 Declare a variable as a ProjectPartCollection. 
            // Instantiate it using the ProjectParts property of the CurrentProject
            // of the PlantProject from step 1.6. 
            ProjectPartCollection projPartCol = currentProj.ProjectParts;


            // 1.8 Use a foreach loop and iterate through the Projects in the 
            // ProjectPartCollection from step 1.7.
            // Note: put the closing curley brace below step 1.17.
            foreach (Project proj in projPartCol)
            {
                // 1.9 Use WriteMessage() of the Editor (ed) created above to print
                // the name of ProjectName of the project. Use a string similar to this 
                // "\nProject Name = " plus the ProjectName property of the Project.
                // (foreach is looping through projects)
                ed.WriteMessage("\nProject Name =  " + proj.ProjectName);


               // 1.10 Declare a string and make it equal to the return from 
               // the ProjectPartName method of the PlantProject from step 1.6
               // Pass in the Project. Name it something like "projPartName".
               string projPartName = currentProj.ProjectPartName(proj);

               // 1.11 Use WriteMessage() of the Editor (ed) created above to print
               // the name of project part. Use a string similar to this 
               // "\nProject Part Name = " plus the string from step 1.10
                ed.WriteMessage("\nProject Part Name = " + projPartName);
             

                // 1.12 Declare a string variable and make it equal to 
                // "\n Project Directory = " plus the ProjectDirectory property 
                // of the Project (used in the foreach loop) 
                string strMsg = "\n Project Directory = " + proj.ProjectDirectory;

                // 1.13 Use the WriteMessage function of the Editor variable from 
                // step 1.5 Pass in the string from step 1.12. This will print the 
                // ProjectDirectory on the commmand line.
                ed.WriteMessage(strMsg);

                // 1.14 Declare a List<> of PnPProjectDrawing named dwgList
                // instantiate it using the GetPnPDrawingFiles() of the project
                List<PnPProjectDrawing> dwgList = proj.GetPnPDrawingFiles();

                // 1.15 Use a foreach loop and iterate through the PnPProjectDrawing 
                // in the List<> from step 1.14 
                // Note: put the closing curley brace below step 1.17
                foreach (PnPProjectDrawing dwg in dwgList)
                {
                    // 1.16 Declare a string variable and make it equal to 
                    // "\n Absolute File name = "  + the AbsoluteFileName of PnPProjectDrawing
                    // variable used in this foreach loop (step 1.16) 
                    strMsg = "\n Absolute File name = " + dwg.AbsoluteFileName;

                    // 1.17 Use the WriteMessage function of the Editor variable from 
                    // step 1.5. Pass in the string from step 1.16 This will print the 
                    // AbsoluteFileName on the commmand line.
                    ed.WriteMessage(strMsg);
                }
            }
        }

        // Project settings are available via PnIdProject, PipingProject classes as methods / properties. 
        // They both share the Project as the base class.
        // Any particular project setting you are interested in?

    
//        The “project database” is a collection of databases, one per project part.

//PlantProject
//                PnIdProject        // processpower.dcf
//                PipingProject     // piping.dcf, 
//                IsoProject           // iso.dcf
//                …                             // etc etc

//In Project.xml, you can see the names that are used:

//So let’s say you direct access to the ProcessPower.dcf database, if you have a PlantProject object:

//      PnPDatabase objDb = objPlantProject.ProjectParts[“PnId”].DataLinksManager.GetPnPDatabase()

//Now, there are several assemblies that must be referenced depending on what you are working with:

//PnPProjectManagerMgd.dll        // PlantProject, Project abstract base class, MiscProject  definition
//PnIdProjectPartsMgd.dll         // PnIdProject; looks to be missing from PlantSDK get it from ACAD.EXE folder
//PnP3dProjectPartsMgd.dll        //  PipingProject, IsoProject
//PnP3dOrthoProjectPart.dll       // OrthoProject




        [CommandMethod("ProjectSettings")]
        public static void Project_Settings()
        {

            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // 1.18 Declare a variable as a PlantProject. Instantiate it using
            // the CurrentProject of PlantApplication. (PlantApplication is a singleton)  
            PlantProject currentProj = PlantApplication.CurrentProject;

            try
            {


                // 1.19 Declare a variable as a PnIdProject. Instantiate it using the
                // ProjectParts property of the CurrentProject of the PlantProject from step 1.18. 
                // Use square brackets [] to specify a string identifier. For the string 
                // property use "PnID". Need to cast it using (PnIdProject)
                // Note: Need to add a reference to PnIdProjectPartsMgd.dll this is dll is Missing
                // from the PlantSDK. Reference it from the ACAD.EXE folder instead.
                PnIdProject PnIdProj = (PnIdProject)currentProj.ProjectParts["PnID"];



                // 1.20 Use the WriteMessage function of the editor created above (ed).
                // Print out the GapWidth property of the PnIdProject from step 1.19.
                // Use a string similar to this: "\nGap Width = " + use ToString()
                ed.WriteMessage("\nGap Width = " + PnIdProj.GapWidth.ToString());


                // 1.21 Change the GapWidth property to .125
                PnIdProj.GapWidth = .125;

                // 1.22 Use the WriteMessage function of the editor created above (ed).
                // Print out the value of the GapLoop property of the PnIdProject from step 1.19.
                // Use a string similar to this: "\nGapLoop  = " + use ToString(). 
                // (this value is a bool)
                ed.WriteMessage("\nGapLoop  = " + PnIdProj.GapLoop().ToString());

              
                // 1.23 Declare a variable as a PipingProject. Instantiate it using the
                // ProjectParts property of the CurrentProject of the PlantProject from step 1.18. 
                // Use square brackets [] to specify a string identifier. For the string 
                // property use "Piping". Need to cast it using (PipingProject)
                // Note: Need to add a reference to PnP3dProjectPartsMgd.dll for PipingProject
                PipingProject pipeProj = (PipingProject)currentProj.ProjectParts["Piping"];

                // 1.24 Use the WriteMessage function of the editor created above (ed).
                // Print out the Minimum Pipe Length property of the PipingProject from step 1.23.
                // Use a string similar to this: "\nMinimum Pipe Length = " + use ToString()
                ed.WriteMessage("\nMinimum Pipe Length = " + pipeProj.MinimumPipeLength.ToString());

                // 1.25 Change the MinimumPipeLength property of the PipingProject from step 1.23.
                // to 10
                pipeProj.MinimumPipeLength = 10;

                // 1.26 Change the WeldGapSize property of the PipingProject from step 1.23.
                // to .125
                pipeProj.WeldGapSize = .125;

                // 1.27 Use the WriteMessage function of the editor created above (ed).
                // Print out the value of UseWeldGaps of the PipingProject from step 1.23.
                // Use a string similar to this: "\nUse Weld Gaps = " + use ToString()
                ed.WriteMessage("\nUse Weld Gaps = " + pipeProj.UseWeldGaps.ToString());

                // 1.28 Change the UseWeldGaps of the PipingProject to true
                pipeProj.UseWeldGaps = true;

            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }

// ***************************** End of LAB 1 ***********************

    }
}
