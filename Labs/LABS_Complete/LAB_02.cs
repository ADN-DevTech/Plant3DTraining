
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Add a reference to PnIDMgd.dll (needed for LineSegment) 
// Add a reference to PnPCommonDbxMgd (needed for LineStyle)

using Autodesk.ProcessPower.PnIDObjects; //(needed for LineSegment)
using Autodesk.ProcessPower.Styles;      // (needed for LineStyle)

namespace LABS_Complete
{
    public class LAB_02
    {
        // ***************** Start of LAB 2 *************************************

        [CommandMethod("LineGroup", CommandFlags.Modal)]
        public static void getLineGroup()
        {

            // Get the TransactionManager
            TransactionManager tm =
            AcadApp.DocumentManager.MdiActiveDocument.Database.TransactionManager;

            // Get the AutoCAD editor
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            //
            PromptEntityOptions pmtEntOpts = new PromptEntityOptions("Select a LineSegment: ");
            PromptEntityResult pmtEntRes = ed.GetEntity(pmtEntOpts);
            if (pmtEntRes.Status == PromptStatus.OK)
            {
                ObjectId entId = pmtEntRes.ObjectId;

                // 2.1 Use the using statement and start a transaction
                // Use the transactionManager created above (tm)
                // Note: Put the closing curly brace after step 2.14
                using (Transaction tr = tm.StartTransaction())
                {

                    // 2.2 Declare a variable as a LineSegment. Instantiate
                    // it by making it equal to the GetObject method of the transaction
                    // from step 2.1. Cast it is a (LineSegment) and open it for read
                    // use the ObjectId from above (entID) for the ObjectId parameter.
                    LineSegment sline = (LineSegment)tm.GetObject(entId, OpenMode.ForRead);

                    // 2.3 Test to see if the LineSegment from step 2.2 is null. If it 
                    // is null return. 
                    if (sline == null) return;

                    // 2.4 Declare a variable as a LineStyle. Instantiate
                    // it by making it equal to the GetOjbect method of the transaction
                    // from step 2.1. Cast it is a (LineStyle). Use the StyleID property
                    // of the LineSegment from step 2.2 and open it for read
                    LineStyle linestyle = (LineStyle)tm.GetObject(sline.StyleID, OpenMode.ForRead, true);

                    // 2.5 Create a string variable make it equal to this:
                    // "\nLineStyle: " +
                    // linestyle.Name; sMsg += "\n   " +
                    // linestyle.FlagType; sMsg += "\n   " +
                    // linestyle.GapPriority; sMsg += "\n   " +
                    // linestyle.ShowFlowDirection ;  sMsg  += "\n";
                    // Note: change "lineStyle" to the variable used in step 2.4 
                    string sMsg = "\nLineStyle: " +
                        linestyle.Name; sMsg += "\n   " +
                        linestyle.FlagType; sMsg += "\n   " +
                        linestyle.GapPriority; sMsg += "\n   " +
                        linestyle.ShowFlowDirection; sMsg += "\n";

                    // 2.6 Use the WriteMessage of the Editor declared above (ed)
                    // to put the string from step 2.5 on the command line
                    // pass in the string
                    ed.WriteMessage(sMsg);


                    // 2.7 Declare a variable as a LineGroupManager and get the line group id
                    // and assign it to variable lineGroupId.
                    LineGroupManager lineGroupMgr = new LineGroupManager();
                    int lineGroupId = lineGroupMgr.GroupId(entId);

                    // 2.8 use the string variable from step 2.5, make it equal to this:
                    // "\nLineGroup: " + lineGroupId.ToString() + " (" + type.ToString() + ")";
                    // Create the type variable first and assign it.
                    GroupType type = lineGroupMgr.Type(lineGroupId);
                    sMsg = "\nLineGroup: " + lineGroupId.ToString() + " (" + type.ToString() + ")";

                    // 2.9 Add to the string variable from step 2.8. (use +=)
                    // add the string "\n   LineSegments: " plus the Count property
                    // of the LineSegments of the LineGroup from step 2.7
                    sMsg += "\n   LineSegments: " + lineGroupMgr.Count(lineGroupId);

                    // 2.10 Use the WriteMessage of the Editor declared above (ed)
                    // to put the string from step 2.8 and 2.9 on the command line
                    // pass in the string
                    ed.WriteMessage(sMsg);

                    // 2.11 use a foreach loop and iterate through the ObjectId of the 
                    // LineSegments of the LineGroup from step 2.7 (use the LineDbIds
                    // property)
                    // Note: Put the closing curly brace after step 2.14
                    foreach (ObjectId lsId in lineGroupMgr.LineDbIds(lineGroupId))
                    {
                        // 2.12 Declare a LineSegment variable and instantiate it using
                        // the transaction from step 1.4. Use the ObjectId from the
                        // loop (step 2.11) Open it for read.
                        LineSegment ls = (LineSegment)tm.GetObject(lsId, OpenMode.ForRead);

                        // 2.13 make the string variable from step 2.5 equal to this string:
                        // "\n   LineSegment: " plus the ClassName property of the LineSegment
                        // from step 2.12 plus this string " ,Vertices 0 - x,y,z "  plus
                        // the Vertices(0).ToString()
                        sMsg = "\n   LineSegment: " + ls.ClassName + " ,Vertices 0 - x,y,z " + ls.Vertices[0].ToString();

                        // 2.14 Use the WriteMessage of the Editor declared above (ed)
                        // to put the string from steps 2.13 the command line
                        // pass in the string
                        ed.WriteMessage(sMsg);

                        // Build and test the code. (run command "LineGroup", you should see
                        // details about the Line Segments on the command line.
                        // Continue to step 2.15 in the lineObjectsAndStylesAccess function

                    }
                }
            }
        }


        [CommandMethod("LineObjects", CommandFlags.Modal)]
        public static void lineObjectsAndStylesAccess()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            Database db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            ObjectId entId = ed.GetEntity("Pick an SLINE: ").ObjectId;
            if (entId.IsNull) return;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;
            using (Transaction tr = tm.StartTransaction())
            {

                Entity entity = (Entity)tm.GetObject(entId, OpenMode.ForRead, true);

                // 2.15 Declare a LineSegment variable make it equal to the entity
                // above "entity" use "as LineSegment" for the cast
                LineSegment sline = entity as LineSegment;

                // 2.16 Use an if statement to see if the LineSegment from step 2.15
                // equals null. If it is null return.
                if (sline == null) return;

                // 2.17 Declare a LineObjectCollection and instantiate it using the 
                // GetLineObjectCollection of the LineSegment from step 2.15
                // Pass in null for the LineObjectFilter argument 
                LineObjectCollection colLineObj = sline.GetLineObjectCollection(null);

                // 2.18 use a foreach statement and iterate through the LineObject
                // in the LineObjectCollection from step 2.17.
                // Note: Put the closing curly brace after step 2.28
                foreach (LineObject lo in colLineObj)
                {
                    // 2.19 Use the WriteMessage method of the editor created above "ed"
                    // use this string "\nSegment " plus the SegmentIndex property of
                    // the LineObject. (from step 2.18 in the foreach). plus this string
                    // ": " plus the GetType method of the LineObject plus this string
                    //  ", " plus the Position property of the LineObject
                    ed.WriteMessage("\nSegment " + lo.SegmentIndex +
                    ": " + lo.GetType() + ", " + lo.Position);

                    // 2.20 The following code would list the ClassName of all Asset items
                    // on the line. Use an if statement and see if the LineObject is
                    // an InLineEntity
                    // Note: Put the closing curly brace below step 2.28
                    if (lo is InlineEntity)
                    {
                        // 2.21 The LineObject is an InLineEntity. Create a variable
                        // as an InLineEntity and make it equal to the LineObject
                        // use (InlineEntity) to cast.
                        InlineEntity loInlineEntity = (InlineEntity)lo;

                        // 2.22 Declare a variable as an entity. Instantiate it using 
                        // GetObject of the transaction created above. (tm). Use the 
                        // ObjectId of the InlineEntity from step 2.21. Open it for read.
                        Entity entityLo = (Entity)tm.GetObject(loInlineEntity.ObjectId, OpenMode.ForRead);

                        // 2.23 Use an if statement and see if the Entity from step 2.22 is 
                        // Asset.
                        // Note: Put the closing curly brace below step 2.28
                        if (entityLo is Asset)
                        {
                            // 2.24 Declare a variable as an Asset. Instantiate it
                            // by making it equal to the Entity from step 2.22 
                            // cast using (Asset)
                            Asset assetLo = (Asset)entityLo;

                            // 2.25 Use the WriteMessage method of the editor created above "ed"
                            // use this string "\nConnected to (" plus the ClassName property of
                            // the Asset from step 2.24 plus this string ")"
                            ed.WriteMessage("\nConnected to (" + assetLo.ClassName + ")");

                            // 2.26 Declare a variable as a Style. Instantiate it using 
                            // GetObject of the transaction created above. (tm). Use the 
                            // StyleId property of the asset (from step 2.24). Open it for read.
                            Style style = (Style)tm.GetObject(assetLo.StyleId, OpenMode.ForRead);

                            // 2.27  Create a string variable make it equal to this:
                            // "\nStyle: " plus the Name property of the Style from 
                            // step 2.13  plus "\n   " repeat this pattern and use
                            // the string variable += to add the ClassName, Decription
                            // and SymbolName of the Style to the string.
                            string sMsg = "\nStyle: " + style.Name;
                            sMsg += "\n   " + style.ClassName;
                            sMsg += "\n   " + style.Description;
                            sMsg += "\n   " + style.SymbolName;


                            // 2.28 Use the WriteMessage method of the editor created
                            // above pass in the string from step 2.27
                            ed.WriteMessage(sMsg);

                        }
                    }
                }
            }
        }

        // ***************** End of LAB 2 *************************************  
    }
}