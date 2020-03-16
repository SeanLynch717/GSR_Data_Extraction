///----------------------------------------------------------------------------------------------------------------------------
/// Author: Sean Lynch
/// Create Date: March 13, 2020
/// Description: Code was developed for a research team at Suny Geneseo.
/// Precondition of data: Data is in an excel file and doesn't start until the 8th line, where we assume time equals zero.
///     We are only interested in the data in column B.Each subsequent data piece is at a time 0.1 seconds from the previous.
///     There are an indeterminate number of data enties, so we rely on a blank cell to indicate when to stop reading.
///----------------------------------------------------------------------------------------------------------------------------
///
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace GSR_Data_Extraction
{
    /// <summary>
    /// Class that handles all the user interaction with the form, as well as processing of the data.
    /// </summary>
    public partial class Form1 : Form
    {
        private string startTimeString;
        private string idString;
        private int startTimeVal;
        private int idVal;
        private int endTime;
        private string fileName;

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheetInput;
        private Excel.Worksheet xlWorkSheetOutput;

        private List<double> inputData;
        private double[,] outputData;

        private bool validId;
        private bool validStartTime;
        private bool opened;

        /// <summary>
        /// Proper constructor for a form object
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            inputData = new List<double>();
        }
        /// <summary>
        /// Runs when the form is set up.
        /// </summary>
        /// <param name="sender"> The form object</param>
        /// <param name="e"> The event args</param>
        private void Form1_Load(object sender, EventArgs e)
        {
            validId = false;
            validStartTime = false;
            opened = false;
            statusRichTextBox.ReadOnly = true;
            statusRichTextBox.Text = "Ready to open excel file.";
            // Set beginning state of the application.
            setActive(openButton);
            setActive(resetButton);
            setInactive(convertButton);
            setInactive(saveButton);
        }
        /// <summary>
        /// Runs when the user clicks the openButton.
        /// The purpose of this method is to process the selected excel file and extract its data.
        /// </summary>
        /// <param name="sender"> The open Button</param>
        /// <param name="e"> The event args</param>
        private void openButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            try
            {
                // Allow only excel files to be opened.
                open.Filter = "Excel Files | *.xlsx";
                open.Title = "Load an excel worksheet";
                DialogResult result = open.ShowDialog();
                // User selected a file.
                if (result == DialogResult.OK)
                {
                    statusRichTextBox.Text = "Reading file. This may take a minute.";
                    // Create a new Excel application
                    xlApp = new Excel.Application();
                    fileName = open.FileName;
                    // Open the existing work book the user selected.
                    xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    // Get a reference to the sheet that contains the data.
                    xlWorkSheetInput = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    string column = "B";
                    int row = 8;
                    int time = 0;
                    // Read data in column B starting at row 8, until there is no more data to read, indicated by a blank cell.
                    while (xlWorkSheetInput.Cells[row, 2].Text != "")
                    {
                        double d = Double.Parse(xlWorkSheetInput.get_Range(column + row, column + row).Value2.ToString());
                        inputData.Add(Double.Parse(xlWorkSheetInput.get_Range(column + row, column + row).Value2.ToString()));
                        time++;
                        row++;
                        //Console.WriteLine(time);
                    }
                    // Overshot by one tenth of a second.
                    endTime = time - 1;
                }
            }
            // Error opening the file.
            catch (Exception oops)
            {
                Console.WriteLine("Error loading file: " + oops.Message);
                statusRichTextBox.Text = "Failed to read file";
            }
            // After try-catch has executed
            finally
            {
                // File was successfully opened.
                if (open.FileName != "")
                {
                    opened = true;
                    // valid id's and start times have already been inputed.
                    if(validStartTime && validId)
                    {
                        // Switch state of the application to be ready to "Convert".
                        setActive(convertButton);                      
                        statusRichTextBox.Text = "Ready to convert excel file.";
                    }
                    // User must still enter valid id and start time.
                    else
                    {
                        statusRichTextBox.Text = "You must enter a valid ID, and start time before converting.";
                    }
                    // Button is no longer active.
                    setInactive(openButton);
                }
            }
        }
        /// <summary>
        /// A trial consists of 120 data entries, or 12 seconds worth of data. For each trial, compte the minimum value from 3 seconds before
        /// the trial to the beginning of the trial.
        /// and the maximum value from the beginning of the trial, and 6 seconds after it begins.
        /// The react is the difference (max - min).
        /// </summary>
        /// <param name="sender"> The convert button</param>
        /// <param name="e"> THe event args</param>
        private void convertButton_Click(object sender, EventArgs e)
        {
            string errorMessage = "";
            // User hasn't entered a valid id.
            if (!validId)
            {
                errorMessage += "Id values must be in the form of integers!\n";
            }
            // User hasn't entered a valid start time.
            if (!validStartTime)
            {
                errorMessage += "Start Time must be in the form of an integer or decimal!\n";
            }
            // There was an issue with at least one input.
            if (!validId || !validStartTime)
            {
                MessageBox.Show(errorMessage, "Error with inputs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            // No issues, proceed to read file.
            else
            {
                statusRichTextBox.Text = "Converting data...";
                // How many trials there are.
                int range = (endTime - startTimeVal) / 120;
                int trial = 1;
                int currentTime = startTimeVal;
                outputData = new double[range, 5];
                // Compute the afformentioned values for at most 100 trials, and at a minimum 'range' number of trials.
                while (trial <= 100 && trial <= range)
                {
                    double min = Double.MaxValue;
                    // March through each data point 3 seconds before the trial has begun, and the beginning of the trial.
                    for (int t = currentTime - 30; t < currentTime; t++)
                    {
                        // We have found a smaller data point.
                        if (inputData[t] < min)
                        {
                            min = inputData[t]; ;
                        }
           
                    }
                    double max = Double.MinValue;
                    // March through each data point from the beginning of the trial to 6 seconds after it starts.
                    for (int t = currentTime; t <= currentTime + 60; t++)
                    {
                        // We have found a larger value.
                        if (inputData[t] > max)
                        {
                            max = inputData[t];
                        }
                    }            
                    double react = max - min;
                    // Update our array to reflect these values.
                    outputData[trial - 1, 0] = idVal;
                    outputData[trial - 1, 1] = trial;
                    outputData[trial - 1, 2] = max;
                    outputData[trial - 1, 3] = min;
                    outputData[trial - 1, 4] = react;
                    // Go to next trial.
                    currentTime += 120;
                    trial++;
                }
                //for (int i = 0; i < range; i++)
                //{
                //    Console.WriteLine(outputData[i, 0] + " " + outputData[i, 1] + " " + outputData[i, 2] + " " + outputData[i, 3] + " " + outputData[i, 4] + " ");
                //}
                
                // Move to the next state of the application: Saving.
                statusRichTextBox.Text = "Ready to save";
                setInactive(convertButton);
                setActive(saveButton);
            }
        }

        /// <summary>
        /// Given that the min's, max's and react's have been calculated, create a new excel sheet and write that information to it.
        /// </summary>
        /// <param name="sender"> The save button</param>
        /// <param name="e"> The event args</param>
        private void saveButton_Click(object sender, EventArgs e)
        {
            // Create a new excel work sheet and store a reference to it.
            xlWorkSheetOutput = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheetOutput.Name = "GSR Output";

            // Set up labels for each column.
            xlWorkSheetOutput.Cells[1, 1] = "ID";
            xlWorkSheetOutput.Cells[1, 2] = "Trial";
            xlWorkSheetOutput.Cells[1, 3] = "Max";
            xlWorkSheetOutput.Cells[1, 4] = "Min";
            xlWorkSheetOutput.Cells[1, 5] = "React";
            // March through each row in our data, and write to our new excel sheet.
            for (int i = 1; i <= outputData.GetLength(0); i++)
            {
                //Console.WriteLine("Added Trial #" + i);
                xlWorkSheetOutput.Cells[i + 1, 1] = outputData[i - 1, 0];
                xlWorkSheetOutput.Cells[i + 1, 2] = outputData[i - 1, 1];
                xlWorkSheetOutput.Cells[i + 1, 3] = outputData[i - 1, 2];
                xlWorkSheetOutput.Cells[i + 1, 4] = outputData[i - 1, 3];
                xlWorkSheetOutput.Cells[i + 1, 5] = outputData[i - 1, 4];
            }
            // Set styles.
            xlWorkSheetOutput.get_Range("A1", "B" + outputData.GetLength(0) + 1).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheetOutput.get_Range("C2", "E" + outputData.GetLength(0) + 1).NumberFormat = "##0.0000";

            try
            {
                SaveFileDialog save = new SaveFileDialog();
                // Can only save as an excel file.
                save.Filter = "Excel Files | *.xlsx";
                save.Title = "Save an excel worksheet";
                DialogResult result = save.ShowDialog();
                // User has specified a name and location to save the file to.
                if (result == DialogResult.OK)
                {
                    // Create new savefile of this excel workbook, with both the original and new worksheet included.
                    xlWorkBook.SaveAs(save.FileName);
                    statusRichTextBox.Text = "File saved. You may begin the process again.";
                    // Reset the application, getting ready to accept a new excel file to process.
                    this.reset();
                }
            }
            // Unable to save file.
            catch (Exception oops)
            {
                Console.WriteLine("Error saving file: " + oops.Message);
            }
        }
        /// <summary>
        /// Release the com object
        /// </summary>
        /// <param name="obj"> The object to release</param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// User has updated the value of idTextBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void idTextBox_TextChanged(object sender, EventArgs e)
        {
            this.idString = idTextBox.Text;
            // Try to parse the value of the text box into an integer.
            try
            {
                idVal = Int32.Parse(idString);
                validId = true;
            }
            // Failes to convert to an integer.
            catch (Exception oops)
            {
                // Not a valid id.
                validId = false;
            }
            // The user has opened a file and both the id and start times are valid...
            if(opened && validId && validStartTime)
            {
                // Transition to next state: Conversion.
                setActive(convertButton);
                statusRichTextBox.Text = "Ready to convert excel file.";
            }
            // Not ready to convert yet.
            else
            {
                setInactive(convertButton);
            }
        }
        /// <summary>
        /// User has updated the value of startTimeTextBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void startTimeTextBox_TextChanged(object sender, EventArgs e)
        {
            this.startTimeString = startTimeTextBox.Text;
            // Try to parse the value of the text box into a double.
            try
            {
                startTimeVal = (int)(Double.Parse(startTimeString) * 10);
                validStartTime = true;
            }
            // Failes to convert to a double.
            catch (Exception oops)
            {
                // Not a valid start time.
                validStartTime = false;
            }
            // The user has opened a file and both the id and start times are valid...
            if (opened && validId && validStartTime)
            {
                // Transition to next state: Conversion.
                setActive(convertButton);
                statusRichTextBox.Text = "Ready to convert excel file.";
            }
            // Not ready to convert yet.
            else
            {
                setInactive(convertButton);
            }
        }
        /// <summary>
        /// Resets the entire application to its starting state.
        /// </summary>
        private void reset()
        {
            if(xlWorkBook != null)
                xlWorkBook.Close(true);
            if(xlApp != null)
                xlApp.Quit();
            if(xlWorkSheetInput != null)
                releaseObject(xlWorkSheetInput);
            if(xlWorkBook != null)
                releaseObject(xlWorkBook);
            if(xlApp != null)
                releaseObject(xlApp);
            xlWorkBook = null;
            xlApp = null;
            xlWorkSheetInput = null;
            xlWorkSheetOutput = null;
            inputData = new List<double>();
            outputData = null;
            startTimeString = null;
            fileName = null;
            opened = false;
            setActive(openButton);
            setInactive(convertButton);
            setInactive(saveButton);
            statusRichTextBox.Text = "Ready to open excel file.";
        }
        /// <summary>
        /// Resets the application when the button is clicked.
        /// </summary>
        /// <param name="sender"> The reset button</param>
        /// <param name="e"> The event args</param>
        private void resetButton_Click(object sender, EventArgs e)
        {
            this.reset();
        }
        /// <summary>
        /// Sets the object to active
        /// </summary>
        /// <param name="sender"></param>
        private void setActive(object sender)
        {
            // Only set to active if it is a button.
            if(sender is Button)
            {
                Button b = (Button)sender;
                b.Enabled = true;
            }
        }
        /// <summary>
        /// Sets the object to inactive
        /// </summary>
        /// <param name="sender"></param>
        private void setInactive(object sender)
        {
            // Only set to inactive if it is a button.
            if (sender is Button)
            {
                Button b = (Button)sender;
                b.Enabled = false;
            }
        }
    }
}
