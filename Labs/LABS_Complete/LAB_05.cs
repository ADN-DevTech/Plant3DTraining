using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.DataLinks; // (needed for DataLinksManager)

namespace LABS_Complete
{
    public class LAB_05
    {
        // ************************** Start of LAB 5 **************************  
        //Events

        [CommandMethod("AddPNPEvent")]
        public void addPNPevent()
        {
            // 5.1 Declare a variable as a PlantProject. Instantiate it using
            // the CurrentProject of PlantApplication
            PlantProject mainPrj = PlantApplication.CurrentProject;

            // 5.2 Declare a Project and instantiate it using  
            // ProjectParts[] of the PlantProject from step 4.1
            // use "PnId" for the name. This will get the P&ID project 
            Project prj = mainPrj.ProjectParts["PnId"];

            // 5.3 Declare a variable as a DataLinksManager. Instantiate it using
            // the DataLinksManager property of the Project from 5.2.
            DataLinksManager dlm = prj.DataLinksManager;

            // 5.4 Add a DataLinkOperationOccurred event. (use the DataLinksManager
            // from step 5.3) Use += and use new to create a new DataLinkEventHandler
            // pass in the name of the function that you will create in step 5.5.
            // (dlm_DataLinkOperationOccurred)
            dlm.DataLinkOperationOccurred += new DataLinkEventHandler(dlm_DataLinkOperationOccurred);
        }

        // 5.5 Create a function named dlm_DataLinkOperationOccurred have it return 
        // void. The name needs to be the same as the name used in step 5.4.
        // It needs two arguments. The first is an object (name it sender). The
        // second argument is a DataLinkEventArgs name it e.
        // Note: Put the closing curly brace below step 5.10
        void dlm_DataLinkOperationOccurred(object sender, DataLinkEventArgs e)
        {
            // 5.7 Instantiate an variable as an Editor using
            // AcadApp.DocumentManager.MdiActiveDocument.Editor; 
            // (AcadApp is the name of the variable from step 1.3)
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // 5.8 Use the WriteMessage function of the Editor from step 5.7)
            // and print the Action of the DataLinkEventArgs passed into this
            // function on the command line. Use a string similar to this:
            // "\nAction = " + use ToString()
            ed.WriteMessage("\nAction = " + e.Action.ToString());

            // 5.9 Use an if statement and see if the RowId is greater than zero
            // Note: Put the closing curly brace below step 5.10
            if (e.RowId > 0)
            {
                // 5.10 Use the WriteMessage function of the Editor from step 5.7)
                // and print the RowId of the DataLinkEventArgs passed into this
                // function on the command line. Use a string similar to this:
                // "\nRowId = " + use ToString()
                ed.WriteMessage("\nRowId = " + e.RowId.ToString());
            }

        }

    }
}