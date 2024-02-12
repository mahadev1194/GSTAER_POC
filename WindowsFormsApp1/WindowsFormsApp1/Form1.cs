//using Gssoft.Gscad.DatabaseServices;
//using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Gssoft.Gscad.DatabaseServices;
//using GstarCAD;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private readonly string selectedLayerName;
        //private string selectedLayers = "";
        public Form1()
        {
            InitializeComponent();
            listBox1.SelectionMode = SelectionMode.MultiSimple;
            //btnFetchLayerDetails.Click += Fetch_Layer_Details;
        }

        private void Search(object sender, EventArgs e)
        {
            // Open file dialog to select .dwg file
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Get the selected file path
                    string filePath = openFileDialog.FileName;

                    // Display the file path in the text box
                    textBox1.Text = filePath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open the file. Original error: " + ex.Message);
                }
            }
        }

        // Event handler for the "Open" button
        private void Open_Files(object sender, EventArgs e)
        {
            // Get the file path from the TextBox
            string filePath = textBox1.Text;

            // Check if the file path is not empty
            if (!string.IsNullOrEmpty(filePath))
            {
                try
                {
                    // Check if GstarCAD is installed
                    string gstarcadPath = @"C:\Program Files\Gstarsoft\GstarCAD2022\gcad.exe";
                    if (File.Exists(gstarcadPath))
                    {
                        // Launch GstarCAD with the selected file
                        Process.Start(gstarcadPath, $"\"{filePath}\"");

                        //SendKeys.SendWait("LAYER\n");

                        // Wait for the LAYER command to execute
                        System.Threading.Thread.Sleep(5000); // Adjust the delay as needed
                    }
                    else
                    {
                        MessageBox.Show("GstarCAD is not installed on this computer.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open the file. Original error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please select a .dwg file first.");
            }
        }
        private dynamic gstarCAD; 
        private readonly List<string> gstarCADLayerNames = new List<string>();

        private void Fetch_Layer_Details(object sender, EventArgs e)
        {
            // Access the GstarCAD application through ActiveX
            //dynamic gstarCAD = null;
            try
            {
                gstarCAD = Marshal.GetActiveObject("GstarCAD.Application");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to access GstarCAD. Error: " + ex.Message);
                return;
            }

            if (gstarCAD != null)
            {
                try
                {
                    gstarCADLayerNames.Clear(); // Clear the previous list of layer names

                    // Retrieve layer names from GstarCAD
                    foreach (dynamic layer in gstarCAD.ActiveDocument.Database.Layers)
                    {
                        string layerName = layer.Name;
                        gstarCADLayerNames.Add(layerName);
                    }

                    // Update the listbox with the fetched layer names
                    UpdateLayerListbox();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to fetch layer details from GstarCAD. Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No document is open in GstarCAD.");
            }
        }

        private void UpdateLayerListbox()
        {
            listBox1.Items.Clear(); // Clear existing items

            // Add fetched layer names to the listbox
            foreach (string layerName in gstarCADLayerNames)
            {
                listBox1.Items.Add(layerName);
            }
        }

        private void Disable_Layer(object sender, EventArgs e)
        {
            // Access the GstarCAD application through ActiveX
            dynamic gstarCAD;
            try
            {
                gstarCAD = Marshal.GetActiveObject("GstarCAD.Application");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to access GstarCAD. Error: " + ex.Message);
                return;
            }

            if (gstarCAD != null)
            {
                try
                {
                    // Access the GstarCAD database
                    Gssoft.Gscad.DatabaseServices.Database db = gstarCAD.ActiveDocument.Database;

                    // Start a transaction
                    using (Gssoft.Gscad.DatabaseServices.Transaction trans = (Transaction)db.TransactionManager.StartTransaction())
                    {
                        // Open the Layer table for read
                        LayerTable acLyrTbl;
                        acLyrTbl = trans.GetObject((ObjectId)db.LayerTableId, OpenMode.ForRead) as LayerTable;

                        foreach (var item in listBox1.SelectedItems)
                        {
                            // Get the name of the selected layer
                            string layerName = item.ToString();

                            // Check if the layer exists in the drawing
                            if (acLyrTbl.Has(layerName))
                            {
                                // Open the LayerTableRecord for write
                                LayerTableRecord acLyrTblRec;
                                acLyrTblRec = trans.GetObject(acLyrTbl[layerName], OpenMode.ForWrite) as LayerTableRecord;

                                // Turn the layer off
                                acLyrTblRec.IsOff = true;
                            }
                            else
                            {
                                MessageBox.Show($"Layer '{layerName}' does not exist in the drawing.");
                            }
                        }

                        // Commit the transaction
                        trans.Commit();

                        MessageBox.Show("Selected layers have been turned off in GstarCAD.");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to disable layers in GstarCAD. Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No document is open in GstarCAD.");
            }
        }


        private void Enable_Layer(object sender, EventArgs e)
        {


            // Access the GstarCAD application through ActiveX
            dynamic gstarCAD;
            try
            {
                gstarCAD = Marshal.GetActiveObject("GstarCAD.Application");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to access GstarCAD. Error: " + ex.Message);
                return;
            }

            if (gstarCAD != null)
            {
                try
                {
                    Gssoft.Gscad.DatabaseServices.Database db = gstarCAD.ActiveDocument.Database;
                    // dynamic db = gstarCAD.ActiveDocument.Database;

                    using (Gssoft.Gscad.DatabaseServices.Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        LayerTable acLyrTbl;
                        acLyrTbl = trans.GetObject(db.LayerTableId,
                                                        OpenMode.ForRead) as LayerTable;
                        foreach (var item in listBox1.SelectedItems)
                        {
                            // Iterate through the layers in the document
                            LayerTableRecord acLyrTblRec = trans.GetObject(acLyrTbl[item.ToString()],
                                             OpenMode.ForWrite) as LayerTableRecord;

                            // Turn the layer off
                            acLyrTblRec.IsOff = false;
                        }

                        // Commit the transaction
                        trans.Commit();

                        MessageBox.Show("Selected layers have been turned on in GstarCAD.");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to disable layers in GstarCAD. Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No document is open in GstarCAD.");
            }
        }

        private void Save_Model(object sender, EventArgs e)
        {
            // Access the GstarCAD application through ActiveX
            dynamic gstarCAD = null;
            try
            {
                gstarCAD = Marshal.GetActiveObject("GstarCAD.Application");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to access GstarCAD. Error: " + ex.Message);
                return;
            }

            if (gstarCAD != null)
            {
                try
                {
                    dynamic activeDocument = gstarCAD.ActiveDocument;

                    if (activeDocument.Modified)
                    {
                        string savePath = @""; 

                        activeDocument.SaveAs(savePath);

                        MessageBox.Show("Model saved successfully.");
                    }
                    else
                    {
                        MessageBox.Show("No changes to save.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to save model. Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No document is open in GstarCAD.");
            }
        }


        private void Close(object sender, EventArgs e)
        {
            this.Close();   
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Access the GstarCAD application through ActiveX
            //dynamic gstarCAD = null;
            try
            {
                gstarCAD = Marshal.GetActiveObject("GstarCAD.Application");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to access GstarCAD. Error: " + ex.Message);
                return;
            }

            if (gstarCAD != null)
            {
                try
                {
                    // Loop through the selected items in the ListBox
                    foreach (int index in listBox1.SelectedIndices)
                    {
                        string selectedLayerName = listBox1.Items[index].ToString();
                        //gstarCAD.ActiveDocument.Database.ExecuteSqlString($"-LAYON {selectedLayerName}\n");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to enable selected layers in GstarCAD. Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No document is open in GstarCAD.");
            }
        }


        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        readonly Process[] processes = Process.GetProcessesByName("GstarCAD");
     
        private void File_Path_Shown(object sender, EventArgs e)
        {

        }
     

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
       
    }
}
