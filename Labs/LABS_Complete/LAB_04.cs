using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.DataLinks; // (needed for DataLinksManager)

// Reference PnP3dProjectPartsMgd.dll  
using Autodesk.ProcessPower.PnP3dObjects; // (needed for Part, ConnectionIterator 

using System; 
using System.Collections.Generic; //Needed for KeyValuePair

namespace LABS_Complete
{
    public class LAB_04
    {
        // ************************** Start of LAB 4 **************************  
        [CommandMethod("PipeWalk")]
        public void pipeWalk()
        {
            // 4.1 Declare a variable as a PlantProject. Instantiate it using
            // the CurrentProject of PlantApplication
            PlantProject mainPrj = PlantApplication.CurrentProject;

            // 4.2 Declare a Project and instantiate it using  
            // ProjectParts[] of the PlantProject from step 4.1
            // use "Piping" for the name. This will get the Piping project 
            Project prj = mainPrj.ProjectParts["Piping"];

            // 4.3 Declare a variable as a DataLinksManager. Instantiate it using
            // the DataLinksManager property of the Project from 4.2.
            DataLinksManager dlm = prj.DataLinksManager;

            //  PipingProject pipingPrj = (PipingProject) mainPrj.ProjectParts["Piping"];
            //  DataLinksManager dlm = pipingPrj.DataLinksManager;


            // Get the TransactionManager
            TransactionManager tm =
            AcadApp.DocumentManager.MdiActiveDocument.Database.TransactionManager;

            // Get the AutoCAD editor
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // Prompt the user to select a pipe entity
            PromptEntityOptions pmtEntOpts = new PromptEntityOptions("Select a Pipe : ");
            PromptEntityResult pmtEntRes = ed.GetEntity(pmtEntOpts);
            if (pmtEntRes.Status == PromptStatus.OK)
            {
                // Get the ObjectId of the selected entity
                ObjectId entId = pmtEntRes.ObjectId;

                // Use the using statement and start a transaction
                // Use the transactionManager created above (tm)
                using (Transaction tr = tm.StartTransaction())
                {

                    try
                    {

                        // 4.4 Declare a variable as a Part. Instantiate it using
                        // the GetObject Method of the Transaction created above (tr)
                        // for the ObjectId argument use the ObjectId from above (entId)
                        // Open it for read. (need to cast it to Part)
                        Part pPart = (Part)tr.GetObject(entId, OpenMode.ForRead);

                        // 4.5 Declare a variable as a PortCollection. Instantiate it 
                        // using the GetPorts method of the Part from step 4.4.
                        // use PortType.All for the PortType.
                        PortCollection portCol = pPart.GetPorts(PortType.All); // (PortType.Both);

                        // 4.6 Use the WriteMessage function of the Editor created above (ed)
                        // print the Count property of the PortCollection from step 4.5.
                        // use a string similar to this: "\n port collection count = "
                        ed.WriteMessage("\n port collection count = " + portCol.Count);

                        // 4.7 Declare a ConnectionManager variable. 
                        // (Autodesk.ProcessPower.PnP3dObjects.ConnectionManager)
                        // Instantiate the ConnectionManager variable by making it
                        // equal to a new Autodesk.ProcessPower.PnP3dObjects.ConnectionManager();
                        ConnectionManager conMgr = new Autodesk.ProcessPower.PnP3dObjects.ConnectionManager();

                        // 4.8 Declare a bool variable named bPartIsConnected and make it false
                        bool bPartIsConnected = false;


                        // 4.9 Use a foreach loop and iterate through all of the Port in
                        // the PortCollection from step 4.5.
                        // Note: Put the closing curly brace below step 4.18
                        foreach (Port pPort in portCol)
                        {
                            // 4.10 Use the WriteMessage function of the Editor created above (ed)
                            // print the Name property of the Port (looping through the ports) 
                            // use a string similar to this: "\nName of this Port = " +
                            ed.WriteMessage("\nName of this Port = " + pPort.Name);

                            // 4.11 Use the WriteMessage function of the Editor created above (ed)
                            // print the X property of the Position from the Port
                            // use a string similar to this: "\nX of this Port = " + 
                            ed.WriteMessage("\nX of this Port = " + pPort.Position.X.ToString());

                            // 4.12 Declare a variable as a Pair and make it equal to a 
                            // new Pair().
                            Pair pair1 = new Pair();

                            // 4.13 Make the ObjectId property of the Pair created in step 4.10 
                            // equal to the ObjectId of the selected Part (entId)
                            pair1.ObjectId = entId;

                            // 4.14 Make the Port property of the Pair created in step 4.10
                            // equal to the port from the foreach cycle (step 4.7)
                            pair1.Port = pPort;


                            // 4.15 Use an if else and the IsConnected method of the ConnectionManager
                            // from step 4.7. Pass in the Pair from step 4.12
                            // Note: Put the else statement below step 4.17 and the closing curly
                            // brace for the else below step 4.18
                            if (conMgr.IsConnected(pair1))
                            {
                                // 4.16 Use the WriteMessage function of the Editor (ed)
                                // and put this on the command line:
                                // "\n Pair is connected "
                                ed.WriteMessage("\n Pair is connected ");

                                // 4.17 Make the bool from step 4.8 equal to true.
                                // This is used in an if statement in step 4.19.
                                bPartIsConnected = true;
                            }
                            else
                            {
                                // 4.18 Use the WriteMessage function of the Editor (ed)
                                // and put this on the command line:
                                // "\n Pair is NOT connected "
                                ed.WriteMessage("\n Pair is NOT connected ");
                            }

                        }


                        // 4.19 Use an If statement and the bool from step 4.8. This will be 
                        // true if one of the pairs tested in loop above loop was connected. 
                        // Note: Put the closing curly brace after step 4.26
                        if (bPartIsConnected)
                        {

                            // 4.20 Declare an ObjectId named curObjID make it 
                            // equal to ObjectId.Null
                            ObjectId curObjId = ObjectId.Null;


                            // 4.21 Declare an int name it rowId 
                            int rowId;

                            // 4.22 Declare a variable as a  ConnectionIterator instantiate it
                            // using the NewIterator method of ConnectionIterator (Autodesk.ProcessPower.PnP3dObjects.)
                            // Use the ObjectId property of the Part from step 4.4 
                            ConnectionIterator connectIter = ConnectionIterator.NewIterator(pPart.ObjectId);                       //need PnP3dObjectsMgd.dll

                            // You could Also use this, need to ensure that pPort is connected
                            // Use the ConnectionManager and a Pair as in the example above.
                            // conIter = ConnectionIterator.NewIterator(pPart.ObjectId, pPort);

                            // 4.23 Use a for loop and loop through the connections in the 
                            // ConnectionIterator from step 4.22. The initializer can be empty.
                            // Use !.Done for the condition. use .Next for the iterator.  
                            // Note: Put the closing curly brace after step 4.26
                            for (; !connectIter.Done(); connectIter.Next())
                            {

                                // 4.24 Make the ObjectId from step 4.20 equal to the ObjectId
                                // property of the ConnectionIterator
                                curObjId = connectIter.ObjectId;

                                // 4.25 Make the integer from step 4.21 equal to the return from 
                                // FindAcPpRowId method of the DataLinksManager from step 4.3.
                                // pass in the ObjectId from step 4.24
                                rowId = dlm.FindAcPpRowId(curObjId);

                                //4.26 Use the WriteMessage function of the Editor (ed)
                                // and pring the integer from step 4.25. Use a string similar to this
                                // this on the command line:
                                // "\n PnId = " +
                                ed.WriteMessage("\n PnId = " + rowId);

                            }

                        }
                    }
                    catch (System.Exception ex)
                    {

                        ed.WriteMessage(ex.ToString());
                    }
                }
            }

        }
        // ************************** END LAB 4 **************************  


        // ***************** Could Add to LAB 4 ******************************

        [CommandMethod("changeDataPipe")]
        public void dataChangePipe()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                PlantProject mainPrj = PlantApplication.CurrentProject;
                Project prj = mainPrj.ProjectParts["Piping"];
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

                    // x.xx Create a new KeyValuePair (System.Collections.Generic)
                    // name it something like oldVal. Make it equal a new KeyValuePair
                    // use string for the Key and Value. (use null for the string key
                    // and string value. 
                    KeyValuePair<string, string> oldValShortDescription =
                        new KeyValuePair<string, string>(null, null);

                    KeyValuePair<string, string> newValShortDescription =
                        new KeyValuePair<string, string>(null, null);

                    KeyValuePair<string, string> oldValPartFamilyLongDesc =
                       new KeyValuePair<string, string>(null, null);

                    KeyValuePair<string, string> newValPartFamilyLongDesc =
                        new KeyValuePair<string, string>(null, null);

                    KeyValuePair<string, string> oldValPartSizeLongDesc =
                       new KeyValuePair<string, string>(null, null);


                    KeyValuePair<string, string> newValPartSizeLongDesc =
                        new KeyValuePair<string, string>(null, null);




                    // Iterate through the entries in the list.
                    for (int i = 0; i < list_Properties.Count; i++)
                    {

                        ed.WriteMessage(list_Properties[i].Key.ToString() + "\n");

                        //areMake the the string for the key from 
                        // step x equal to the Key property of the list in
                        // this iteration of the loop "[i]". (The list form step 3.14
                        //string strKey = list_Properties[i].Key;

                        // x.xx Use and if statement and see if the Key property
                        // of the list in this iteration of the loop "[i]" is 
                        // equal to "Description"
                        // Note: put the closing curly brace after step 3.25
                        if (list_Properties[i].Key == "ShortDescription") //Try "ShortDescription" "PartFamilyLongDesc" "PartSizeLongDesc"
                        {

                            // x.xx Make the KeyValuePair created in step 3.20 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property of the List in this 
                            // iteration of the loop. 
                            oldValShortDescription = new KeyValuePair<string, string>(list_Properties[i].Key,
                                       list_Properties[i].Value);

                            // 3.xx Declare a string variable and make it equal to something
                            // like "Some new Manufacturer"
                            string txtNewVal = "Some new Short Description";

                            // x.xx Make the KeyValuePair created in step 3.21 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property use the string
                            // from step 3.24
                            newValShortDescription = new KeyValuePair<string, string>(list_Properties[i].Key,
                                txtNewVal);


                            // x.xx exit the for loop by adding break.
                            //break;
                            continue;

                        }

                        if (list_Properties[i].Key == "PartFamilyLongDesc") //Try "ShortDescription" "PartFamilyLongDesc" "PartSizeLongDesc"
                        {

                            // x.xx Make the KeyValuePair created in step 3.20 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property of the List in this 
                            // iteration of the loop. 
                            oldValPartFamilyLongDesc = new KeyValuePair<string, string>(list_Properties[i].Key,
                                       list_Properties[i].Value);

                            // x.xx Declare a string variable and make it equal to something
                            // like "Some new Manufacturer"
                            string txtNewVal = "Some new Long Description";

                            // x.xx Make the KeyValuePair created in step 3.21 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property use the string
                            // from step 3.24
                            newValPartFamilyLongDesc = new KeyValuePair<string, string>(list_Properties[i].Key,
                                txtNewVal);


                            // x.xx exit the for loop by adding break.
                            //break;
                            continue;

                        }
                        if (list_Properties[i].Key == "PartSizeLongDesc") //Try "ShortDescription" "PartFamilyLongDesc" "PartSizeLongDesc"
                        {

                            // x.xx Make the KeyValuePair created in step 3.20 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property of the List in this 
                            // iteration of the loop. 
                            oldValPartSizeLongDesc = new KeyValuePair<string, string>(list_Properties[i].Key,
                                       list_Properties[i].Value);

                            // x.xx Declare a string variable and make it equal to something
                            // like "Some new Manufacturer"
                            string txtNewVal = "Some new Size long Description";

                            // x.xx Make the KeyValuePair created in step 3.21 equal to
                            // a new KeyValuePair. Use string for the Key and Value. For the 
                            // Key, use Key property of the List in this iteration of the loop
                            // [i]. For the Value use the Value property use the string
                            // from step 3.24
                            newValPartSizeLongDesc = new KeyValuePair<string, string>(list_Properties[i].Key,
                                txtNewVal);

                            // x.xx exit the for loop by adding break.
                            // break;
                            continue;

                        }
                    }


                    // x Remove the old KeyValuePair from the List by calling the 
                    // Remove method from the list created above. (list_Properties) Pass
                    // in the old KeyValuePair from step 3.20
                    list_Properties.Remove(oldValShortDescription);
                    list_Properties.Remove(oldValPartFamilyLongDesc);
                    list_Properties.Remove(oldValPartSizeLongDesc);

                    // x Add the new KeyValuePair to the List by calling the 
                    // Add method from the list created above. (list_Properties) Pass
                    // in the KeyValuePair from step 3.21
                    list_Properties.Add(newValShortDescription);
                    list_Properties.Add(newValPartFamilyLongDesc);
                    list_Properties.Add(newValPartSizeLongDesc);

                    // x Declare a new System.Collections.Specialized.StringCollection
                    // name it something like strNames. Instantiate it by making it 
                    // equal to a new System.Collections.Specialized.StringCollection()
                    System.Collections.Specialized.StringCollection strNames =
                        new System.Collections.Specialized.StringCollection();

                    // x Declare a new System.Collections.Specialized.StringCollection
                    // name it something like strValues. Instantiate it by making it 
                    // equal to a new System.Collections.Specialized.StringCollection()
                    System.Collections.Specialized.StringCollection strValues =
                        new System.Collections.Specialized.StringCollection();

                    // x Iterate through the List declared above (list_Properties)
                    // Note: Put the closing curly brace below step 3.34
                    for (int i = 0; i < list_Properties.Count; i++)
                    {
                        // x Declare a string named something like "name" make
                        // it equal to the Key Property of this Iteration of the list 
                        // in the loop. "[i]"
                        String name = list_Properties[i].Key;

                        // x Declare a string named something like "value" make
                        // it equal to the Value property of this Iteration of the list 
                        // in the loop. "[i]"
                        String value = list_Properties[i].Value;

                        // x Use the Add method of the StringCollection created in 
                        // step 3.28. Pass in the string from step 3.31
                        strNames.Add(name);


                        // x Use the Add method of the StringCollection created in 
                        // step 3.29. Pass in the string from step 3.32
                        strValues.Add(value);
                    }

                    // x Use the SetProperties method of the DataLinksManager created
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

    }
}