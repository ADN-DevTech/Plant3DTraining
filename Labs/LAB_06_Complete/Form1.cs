using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

// ************************* Start of LAB 6 ********************************************
// 6.1 Add the following References From the SDK
// PnPDataObjects.dll  - Needed for Autodesk.ProcessPower.DataObjects
// PnPSQLiteEngine.dll - Needed to access the tables

// 6.2 Add this reference from the Acad.exe directory 
// PnPCommonMgd.dll - for Autodesk.ProcessPower.Common.PnPGenericNamedCollection`1<T0>


// 6.3 Add this NameSpace with a using statement
using Autodesk.ProcessPower.DataObjects;



namespace LAB_06_Complete
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 6.4 Declare a string variable named someDcfFileName. Make
            // it equal to the path of the ProcessPower.dcf of one of
            // the projects on your system. (you can use the @"" to avoid
            // having to use two back slashes in the path.
            string someDcfFileName = @"C:\delme\ProcessPower.dcf";
           
            // 6.5 Decalre a string named tableName make it equal to 
            // an empty string ""
            string tableName = ""; 

            foreach (Control crtl in this.groupBox1.Controls)
            {
                RadioButton radioB = (RadioButton)crtl;
                if( radioB.Checked)
                {
                    // 6.6 Make the string created in step 6.2
                    // equal to the Text property of the radio
                    // button. radioB.
                    tableName = radioB.Text;
                }
            }

            // 6.7 call the function that you create in step 6.8
            // Suggested name is getData. Pass in the string from step 
            // 6.1 and the string from 6.2
            getData(someDcfFileName, tableName);
        }
        
        
        // 6.8 Create a private void function named getData.
        // have it take two string parameters. Name one of the 
        // parameters dcfFileName and the other tableName
        // Note: Put the closing curly brace after step 6.19
        private void getData(string dcfFileName, string tableName)
        {

            // 6.9 Call the Clear method of the Items
            // property of the ListBox on the form. (listBox1)
            listBox1.Items.Clear();
           
            // 6.10 Use the Add method of the Items collection
            // of the ListBox. (listBox1). to add the string
            // passed into the function. (tableName) Use a string
            // similar to this: "Table Name: " +
            listBox1.Items.Add("Table Name: " + tableName);

            // 6.11 Declare a variable as a PnPDatabase. Instantiate
            // it using the Open method of PnPDatabase. Pass in
            // the dcfFileName string passed into this function. 
            PnPDatabase db = PnPDatabase.Open(dcfFileName);
            
            // 6.12 Declare a variable as a PnPTable instantiate it
            // using the Tables collection of the PnPDatabase from step 6.11
            // Pass in the tableName string that was passed into this 
            // function. 
            PnPTable tbl = db.Tables[tableName];

            // 6.13 Declare an array of PnPRow. ( PnPRow[]) Name it
            // rows and make it equal to the Select method of the 
            // PnPTable from step 6.12
            PnPRow[] rows = tbl.Select();

            // 6.14 use a foreach and iterate through each PnPRow
            // in the PnPRow array from step 6.13
            // Note: Put the closing curly brace after step 6.19 
            foreach (PnPRow row in rows)
            {
                // 6.15. Use another foreach and iterate throuh each PnPColumn
                // of the PnPRow from the foreach in step 6.14). 
                // (use the Table.Columns collection) 
                // Note: Put the closing curly brace after step 6.17
                foreach (PnPColumn col in row.Table.Columns)
                {
                   
                    // 6.16 Use the Add method of the Items collection
                    // of the ListBox. (listBox1) to add the Name property
                    // of the PnPColumn in the foreach loop. (step 6.15)
                    listBox1.Items.Add("Column Name = " + col.Name);

                    // 6.17 Use the Add method of the Items collection
                    // of the ListBox. (listBox1) to add the text from the row
                    // in this column. Use the row from the foreach in step 6.14
                    // Use square brackets (row[]) and the Name property of the PnPColumn
                    // from the foreach in step 6.15. (string for the Column) 
                    // Use a string similar to this: "Row Value in this column = " +
                    listBox1.Items.Add("Row Value in this column = " + row[col.Name]);
  
                }
                
                // 6.18 Use the Add method of the Items collection
                // of the ListBox. (listBox1) to add the an empty string
                // this will put a blank line between the rows 
                listBox1.Items.Add(" ");

                // 6.19 Use the Add method of the Items collection
                // of the ListBox. (listBox1) to add the a string "NEXT ROW"
                // This will make the text in the list box easier to read
                listBox1.Items.Add("NEXT ROW ");
            }
        }
    }
}
