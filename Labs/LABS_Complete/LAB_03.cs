

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.DataLinks; // (needed for DataLinksManager)

using Autodesk.ProcessPower.PnPCommonDbx; //FormatStringUtils

using System;
using System.Collections.Generic;
//using System.Text;
//using System.Linq;

namespace LABS_Complete
{
    public class LAB_03
    {
        // ***************** Start of LAB 3 *************************************  

       
        // DataLinksManager.FindAcPpObjectIds(rowId) --> collection of PpObjectId 
        // DataLinksManager.MakeAcDbObjectId(ppObjId) --> converts PpObjectId to ObjectId 
        // There are various flavors of those functions. The main point to 
        // get across is that multiple AcDbObjectIds may be linked to a single RowID.

        [CommandMethod("ConvertIds")]
        public static void convertIds()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            try
            {

                // 3.1 Declare a variable as a PlantProject and instantiate
                // it using the CurrentProject of PlantApplication.
                PlantProject mainPrj = PlantApplication.CurrentProject;

                // 3.2 Declare a variable as a Project and instantiate it
                // using the ProjectParts property of the PlantProject from
                // step 3.1. for the strName property [] use "PnId"
                Project prj = mainPrj.ProjectParts["PnId"];


                //wb testing, commented out the line above
               // Project prj = mainPrj.ProjectParts["Piping"];

                // 3.3 Declare a DataLinksManager variable and instantiate it
                // using the DataLinksManager property of the Project from 
                // step 3.2
                DataLinksManager dlm = prj.DataLinksManager;

                // 3.4 Declare a variable as an ObjectId. Instantiate it
                // use the GetEntity method of the Editor from above. (ed)
                // Use the ObjectId property.
                ObjectId entId = ed.GetEntity("Pick a P&ID item: ").ObjectId;

                // 3.5 Get the Row Id of the entity from the ObjectId
                // from step 3.4. Declare an int and make it equal to the 
                // return from the FindAcPpRowId of the DataLinksManager  from 
                // step 3.3. pass in the ObjectId from step 3.4.
                // Note: The int for the Row Id from this step and the Row Id from
                // step 7 will be the same, even though we are using ObjectId here
                // and PpObjectId in step 3.7
                int rowId1 = dlm.FindAcPpRowId(entId);

                // 3.6 Declare a variable as a PpObjectId and instantiate it
                // using the MakeAcPpObjectId method of the DataLinksManager
                // from step 3.3. Pass in the ObjectId from step 3.4.
                PpObjectId ppObjId = dlm.MakeAcPpObjectId(entId);

                // 3.7 Now we will get the Row Id of the entity from the PpObjectId
                // from step 3.5. Declare an int and make it equal to the 
                // return from the FindAcPpRowId of the DataLinksManager  from 
                // step 3.3. Pass in the PpObjectId 3.6
                int rowId2 = dlm.FindAcPpRowId(ppObjId);

                // 3.8 Use the WriteMessage function of the editor (ed) and print
                // the values of the int from step 3.7 and 3.5
                ed.WriteMessage("rowId1 = " + rowId1.ToString() + " rowId2 = " + rowId2);


                // 3.9 Declare a variable as a PpObjectIdArray. Instantiate it
                // using the FindAcPpObjectIds method of the DataLinksManager from
                // step 3.3. Pass in the row id from step or 3.5. (or step 3.7)
                // NOTE: FindAcPpObjectIds returns a COLLECTION of AcPpObjectId.
                // I.e., multiple AcDbObjectIds may be linked to a single RowID
                PpObjectIdArray ids = dlm.FindAcPpObjectIds(rowId1);

                // 3.10 Use a foreach and iterate through the PpObjectId  in the 
                // PpObjectIdArray from step 3.9. 
                // Note: put the closing curly brace below Step 3.12
                foreach (PpObjectId ppid in ids)
                {
                    // 3.11 Declare a variable as an ObjectId Instantiate it using
                    // the MakeAcDbObjectId variable of the DataLinksManager from
                    // step 3.3.
                    ObjectId oid = dlm.MakeAcDbObjectId(ppid);

                    // 3.12 Use the WriteMessage function of the editor (ed) and print
                    // the value of the ObjectId form step 3.11 
                    // Build, debug and test this code. (Command "ConvertIds")
                    // Note: Continue to step 3.13 below
                    ed.WriteMessage("\n oid = " + oid.ToString() + "\n");

                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }

        }

        [CommandMethod("GetDataByMGR")]
        public void getDataUsingManager()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                PlantProject mainPrj = PlantApplication.CurrentProject;
                Project prj = mainPrj.ProjectParts["PnId"];
                DataLinksManager dlm = prj.DataLinksManager;

                PromptIntegerOptions pOpts = new PromptIntegerOptions("Enter a PnPID number");
                PromptIntegerResult prIntRes = ed.GetInteger(pOpts);

                if (prIntRes.Status == PromptStatus.OK)
                {

                    // 3.13 Declare a List (System.Collections.Generic) Use KeyValuePair
                    // with string for the two elements in the List. (Key, Value)
                    List<KeyValuePair<string, string>> list_Properties;


                    // 3.14 Instantiate the List from step 3.13. Use the 
                    // GetAllProperties method of the DataLinksManager 
                    // from above (dlm). For the RowId argument using the 
                    // Value property of the PromptIntegerResult from above.
                    // (prIntRes) Use true for the current version property.
                    list_Properties = dlm.GetAllProperties(prIntRes.Value, true);

                    // 3.15 Declare a couple of strings. These will be used to get
                    // the Key and the Value from the entries in the list. Name them
                    // something like strKey and strValue.
                    string strKey, strValue = null;

                    // 3.16 Iterate through the entries in the list. 
                    // Use a for statement. Something like the example below. (Change
                    // list_Properties to the name of the List from step 3.14.
                    // for (int i = 0; i < list_Properties.Count; i++)
                    // Note: put the closing curly brace below step 3.19
                    for (int i = 0; i < list_Properties.Count; i++)
                    {
                        // 3.17 Make the string for the key from 
                        // step 3.15 equal to the Key property of the list in
                        // this iteration of the loop "[i]". (The list form step 3.14)
                        strKey = list_Properties[i].Key;

                        // 3.18 Make the string for the value from step 3.15
                        // equal to the Value property of the list in
                        // this iteration of the loop "[i]" of the list form step 3.14
                        strValue = list_Properties[i].Value;

                        // 3.19 Use the WriteMessage function of the editor (ed) and print
                        // the values of the string with the key and the string with the
                        // the value (from steps 3.17 and 3.18) Use "\n" for a return. 
                        // (add it to the end + "\n")
                        // Build, debug and test this code. (Command "GetDataByMGR")
                        // Note: Continue to step 3.20 below
                        ed.WriteMessage("Key = " + strKey + " Value = " + strValue + "\n");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }


        [CommandMethod("changeData")]
        public void dataChange()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                PlantProject mainPrj = PlantApplication.CurrentProject;
                Project prj = mainPrj.ProjectParts["PnId"];
                DataLinksManager dlm = prj.DataLinksManager;

                PromptIntegerOptions pOpts = new PromptIntegerOptions("Enter a PnPID number");
                PromptIntegerResult prIntRes = ed.GetInteger(pOpts);

                int iRowId = prIntRes.Value;

                if (prIntRes.Status == PromptStatus.OK)
                {
                    // Declare a List (System.Collections.Generic) of the properties
                    // from the selected entity
                    List<KeyValuePair<string, string>> list_Properties;
                    list_Properties = dlm.GetAllProperties(iRowId, true);

                    // 3.20 Create a new KeyValuePair (System.Collections.Generic)
                    // name it something like oldVal. Make it equal a new KeyValuePair
                    // use string for the Key and Value. (use null for the string key
                    // and string value. 
                    KeyValuePair<string, string> oldVal =
                        new KeyValuePair<string, string>(null, null);

                    // 3.21 Create another KeyValuePair (System.Collections.Generic)
                    // name it something like newVal. Make it equal a new KeyValuePair
                    // use string for the Key and Value. (use null for the string key
                    // and string value. 
                    KeyValuePair<string, string> newVal =
                        new KeyValuePair<string, string>(null, null);


                    // Iterate through the entries in the list.
                    for (int i = 0; i < list_Properties.Count; i++)
                    {

                        //areMake the the string for the key from 
                        // step 3.15 equal to the Key property of the list in
                        // this iteration of the loop "[i]". (The list form step 3.14
                        //string strKey = list_Properties[i].Key;

                        // 3.22 Use and if statement and see if the Key property
                        // of the list in this iteration of the loop "[i]" is 
                        // equal to "Manufacturer"
                        // Note: put the closing curly brace after step 3.25
                        if (list_Properties[i].Key == "Manufacturer")
                        {

                            // 3.23 Make the KeyValuePair created in step 3.20 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property of the List in this 
                            // iteration of the loop. 
                            oldVal = new KeyValuePair<string, string>(list_Properties[i].Key,
                                       list_Properties[i].Value);

                            // 3.24 Declare a string variable and make it equal to something
                            // like "Some new Manufacturer"
                            string txtNewVal = "Some new Manufacturer";

                            // 3.25 Make the KeyValuePair created in step 3.21 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property use the string
                            // from step 3.24
                            newVal = new KeyValuePair<string, string>(list_Properties[i].Key,
                                txtNewVal);

                            // 3.25 exit the for loop by adding break.
                            break;

                        }
                    }


                    // 3.26 Remove the old KeyValuePair from the List by calling the 
                    // Remove method from the list created above. (list_Properties) Pass
                    // in the old KeyValuePair from step 3.20
                    list_Properties.Remove(oldVal);

                    // 3.27 Add the new KeyValuePair to the List by calling the 
                    // Add method from the list created above. (list_Properties) Pass
                    // in the KeyValuePair from step 3.21
                    list_Properties.Add(newVal);

                    // 3.28 Declare a new System.Collections.Specialized.StringCollection
                    // name it something like strNames. Instantiate it by making it 
                    // equal to a new System.Collections.Specialized.StringCollection()
                    System.Collections.Specialized.StringCollection strNames =
                        new System.Collections.Specialized.StringCollection();

                    // 3.29 Declare a new System.Collections.Specialized.StringCollection
                    // name it something like strValues. Instantiate it by making it 
                    // equal to a new System.Collections.Specialized.StringCollection()
                    System.Collections.Specialized.StringCollection strValues =
                        new System.Collections.Specialized.StringCollection();

                    // 3.30 Iterate through the List declared above (list_Properties)
                    // Note: Put the closing curly brace below step 3.34
                    for (int i = 0; i < list_Properties.Count; i++)
                    {
                        // 3.31 Declare a string named something like "name" make
                        // it equal to the Key Property of this Iteration of the list 
                        // in the loop. "[i]"
                        String name = list_Properties[i].Key;

                        // 3.32 Declare a string named something like "value" make
                        // it equal to the Value property of this Iteration of the list 
                        // in the loop. "[i]"
                        String value = list_Properties[i].Value;

                        // 3.33 Use the Add method of the StringCollection created in 
                        // step 3.28. Pass in the string from step 3.31
                        strNames.Add(name);


                        // 3.34 Use the Add method of the StringCollection created in 
                        // step 3.29. Pass in the string from step 3.32
                        strValues.Add(value);
                    }

                    // 3.35 Use the SetProperties method of the DataLinksManager created
                    // above (dlm) to update the properties. Pass in the Row Id that was
                    // provided by the user above. (iRowId) For the first StringCollection
                    // pass in the one from step 3.28. For the second StringCollection
                    // pass in the one from step 3.29. 
                    dlm.SetProperties(iRowId, strNames, strValues);


                    // Build debug and test this code. In the DataManager
                    // You should see a new value for the Manufacturer.
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }




        // ************************** END LAB 3 **************************  

        
        // ***************** Could Add to LAB 3******************************
        [CommandMethod("GetDataByString")]
        public static void getDataUsingString()
        {

            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                PlantProject mainPrj = PlantApplication.CurrentProject;
                Project prj = mainPrj.ProjectParts["PnId"];
                DataLinksManager dlm = prj.DataLinksManager;

                ObjectId entId = ed.GetEntity("Pick a P&ID item: ").ObjectId;

                // x.x Declare a variable as a PpObjectId and instantiate it
                // using the MakeAcPpObjectId method of the DataLinksManager
                // from step 3.x. Pass in the ObjectId from step 3.4.
                PpObjectId pnpId = dlm.MakeAcPpObjectId(entId);

                // x.x Now let's do an opposite action that we did in step 3.5
                // Now we will get the ObjectId(s) of the entity from the PpObjectId
                // from step 3.x. Declare an int and make it equal to the 
                // return from the FindAcPpRowId of the DataLinksManager  from 
                // step 3.x.
                int rowId1 = dlm.FindAcPpRowId(entId); // You can use ObjectId
                int rowId2 = dlm.FindAcPpRowId(pnpId); //          or PpObjectId
                // rowId1 and rowId2 are always equal

                PpObjectIdArray ids = dlm.FindAcPpObjectIds(rowId1);
                // NOTE: It returns a COLLECTION of AcPpObjectId!
                //       I.e., multiple AcDbObjectIds may be linked to a single RowID

                // Now find the ObjectID for each PpObjectId
                foreach (PpObjectId ppid in ids)
                {
                    ObjectId oid = dlm.MakeAcDbObjectId(ppid);
                    ed.WriteMessage("\n oid=" + oid.ToString() + "\n");

                    // Evaluate the next two lines are not in the DevNote:
                    String sEval = "\n LineNumber = ";

                    sEval += FormatStringUtils.Evaluate
                        ("#(TargetObject.LineNumber^@NNN)", oid); // - #(Project.General.Project_Name)", oid);

                    sEval += "\n Project_Name = ";
                    sEval += FormatStringUtils.Evaluate("#(Project.General.Project_Name)", oid);


                    ed.WriteMessage(sEval);

                    // String sEval2 = "\n TargeObject.Tag = ";
                    // String sTestIsValid = "#(TargetObject.Tag) - #(=TargetObject.Tag)";
                    //String sTestIsValid = "#(TargetObject.Tag)"; // - #(=TargetObject.Tag)";
                    String sTestIsValid = "#(GenericRotaryValve.Manufacturer)";

                    if (FormatStringUtils.IsValid(sTestIsValid))
                    {
                        String sEval2 = "\n Generic Rotary Valve Manufacturer = ";
                        // sEval2 += FormatStringUtils.Evaluate("#(TargetObject.Tag) - #(=TargetObject.Tag)");
                        sEval2 += FormatStringUtils.Evaluate(sTestIsValid, oid);
                        ed.WriteMessage(sEval2);

                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }

        [CommandMethod("FormatStrings", CommandFlags.Modal)]
        public static void formatStringsAndAnnotations()
        {

            //Database db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            //// DataLinksManager dlm = DataLinksManager.GetManager(db);
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            //ObjectId entId = ed.GetEntity("Pick an SLINE TargetObject: ").ObjectId;

            ////Evaluate

            //String sEval = Autodesk.ProcessPower.PnPCommonDbx.FormatStringUtils.Evaluate(
            //   "#(TargetObject.LineNumber^@NNN) - #(Project.General.Project_Name)",
            //   entId
            //);

            //ed.WriteMessage(sEval);

            // Need to figure out how to get the Annotation Class and 
            // AnnotationFormatToolsUtil - This is out of date!
            // use FormatStringUtils instead.
            ObjectId annotationId = ed.GetEntity("Pick Annotation: ").ObjectId;

            Annotation an = new Annotation(annotationId);

            for (int ifs = 0; ifs < an.NumFormatStrings; ifs++)
            {
                String sfs = an.GetFormatString(ifs);
                ObjectId targetId = an.TargetId;
                // AnnotationFormatToolsUtil oUtil = new AnnotationFormatToolsUtil();
                //  FormatStringUtils oUtil = new FormatStringUtils();

                //Evaluate
                // String seval = oUtil.Evaluate(sfs, targetId); //WB commented
                String seval = FormatStringUtils.Evaluate(sfs, targetId); //WB added
                ed.WriteMessage(seval + " (" + sfs + ").");
            }
        }

    }
}