/*
 * @Author: Sean Lynch
 * Developed for educational purp
 */
 
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


        public Form1()
        {
            InitializeComponent();
            inputData = new List<double>();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            validId = false;
            validStartTime = false;
            opened = false;
            statusRichTextBox.ReadOnly = true;
            statusRichTextBox.Text = "Ready to open excel file.";
            setActive(openButton);
            setActive(resetButton);
            setInactive(convertButton);
            setInactive(saveButton);
        }

        private void openButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            try
            {
                open.Filter = "Excel Files | *.xlsx";
                open.Title = "Load an excel worksheet";
                DialogResult result = open.ShowDialog();
                if (result == DialogResult.OK)
                {
                    statusRichTextBox.Text = "Reading file. This may take a minute.";
                    xlApp = new Excel.Application();
                    fileName = open.FileName;
                    xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheetInput = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    string column = "B";
                    int row = 8;
                    int time = 0;
                    // Read data into a dictionary where the time is the key, and the data is the value.
                    while (xlWorkSheetInput.Cells[row, 2].Text != "")
                    {
                        double d = Double.Parse(xlWorkSheetInput.get_Range(column + row, column + row).Value2.ToString());
                        inputData.Add(Double.Parse(xlWorkSheetInput.get_Range(column + row, column + row).Value2.ToString()));
                        time++;
                        row++;
                        Console.WriteLine(time);
                    }
                    // Overshot by one tenth of a second.
                    endTime = time - 1;
                }
            }
            catch (Exception oops)
            {
                Console.WriteLine("Error loading file: " + oops.Message);
                statusRichTextBox.Text = "Failed to read file";
            }
            finally
            {
                if (open.FileName != "")
                {
                    opened = true;
                    if(validStartTime && validId)
                    {
                        setActive(convertButton);                      
                        statusRichTextBox.Text = "Ready to convert excel file.";
                    }
                    else
                    {
                        statusRichTextBox.Text = "You must enter a valid ID, and start time before converting.";
                    }
                    setInactive(openButton);
                }
            }
        }
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

        private void convertButton_Click(object sender, EventArgs e)
        {
            string errorMessage = "";
            if (!validId)
            {
                errorMessage += "Id values must be in the form of integers!\n";
            }
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
                int range = (endTime - startTimeVal) / 120;
                int trial = 1;
                int currentTime = startTimeVal;
                outputData = new double[range, 5];
                Console.WriteLine();
                while (trial <= 100 && trial <= range)
                {
                    double min = Double.MaxValue;
                    for (int t = currentTime - 30; t < currentTime; t++)
                    {
                        if (inputData[t] < min)
                        {
                            min = inputData[t]; ;
                        }
           
                    }
                    double max = Double.MinValue;
                    for (int t = currentTime; t <= currentTime + 60; t++)
                    {
                        if (inputData[t] > max)
                        {
                            max = inputData[t];
                        }
                    }            
                    double react = max - min;
                    outputData[trial - 1, 0] = idVal;
                    outputData[trial - 1, 1] = trial;
                    outputData[trial - 1, 2] = max;
                    outputData[trial - 1, 3] = min;
                    outputData[trial - 1, 4] = react;
                    currentTime += 120;
                    trial++;
                }
                for (int i = 0; i < range; i++)
                {
                    Console.WriteLine(outputData[i, 0] + " " + outputData[i, 1] + " " + outputData[i, 2] + " " + outputData[i, 3] + " " + outputData[i, 4] + " ");
                }
                statusRichTextBox.Text = "Ready to save";
                setInactive(convertButton);
                setActive(saveButton);
            }
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            try
            {
                xlWorkSheetOutput = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheetOutput.Name = "GSR Output";

                xlWorkSheetOutput.Cells[1, 1] = "ID";
                xlWorkSheetOutput.Cells[1, 2] = "Trial";
                xlWorkSheetOutput.Cells[1, 3] = "Max";
                xlWorkSheetOutput.Cells[1, 4] = "Min";
                xlWorkSheetOutput.Cells[1, 5] = "React";
                for (int i = 1; i <= outputData.GetLength(0); i++)
                {
                    //Console.WriteLine("Added Trial #" + i);
                   //Console.WriteLine(String.Format("{0:0.0000}", outputData[i - 1, 2]));
                    xlWorkSheetOutput.Cells[i + 1, 1] = outputData[i - 1, 0];
                    xlWorkSheetOutput.Cells[i + 1, 2] = outputData[i - 1, 1];
                    xlWorkSheetOutput.Cells[i + 1, 3] = outputData[i - 1, 2];
                    xlWorkSheetOutput.Cells[i + 1, 4] = outputData[i - 1, 3];
                    xlWorkSheetOutput.Cells[i + 1, 5] = outputData[i - 1, 4];
                }
                xlWorkSheetOutput.get_Range("A1", "B" + outputData.GetLength(0) + 1).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheetOutput.get_Range("C2", "E" + outputData.GetLength(0) + 1).NumberFormat = "##0.0000";
            }
            catch (Exception oops)
            {
                Console.WriteLine("Error saving file: " + oops.Message);
            }
            finally
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "Excel Files | *.xlsx";
                save.Title = "Save an excel worksheet";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //xlWorkBook.SaveCopyAs(fileName);
                    xlWorkBook.SaveAs(save.FileName);
                    statusRichTextBox.Text = "File saved. You may begin the process again.";
                    this.reset();
                }            
            }
        }

        private void idTextBox_TextChanged(object sender, EventArgs e)
        {
            this.idString = idTextBox.Text;
            try
            {
                idVal = Int32.Parse(idString);
                validId = true;
            }
            catch (Exception oops)
            {
                validId = false;
            }

            if(opened && validId && validStartTime)
            {
                setActive(convertButton);
                statusRichTextBox.Text = "Ready to convert excel file.";
            }
            else
            {
                setInactive(convertButton);
            }
        }

        private void startTimeTextBox_TextChanged(object sender, EventArgs e)
        {
            this.startTimeString = startTimeTextBox.Text;
            try
            {
                startTimeVal = (int)(Double.Parse(startTimeString) * 10);
                validStartTime = true;
            }
            catch (Exception oops)
            {
                validStartTime = false;
            }

            if (opened && validId && validStartTime)
            {
                setActive(convertButton);
                statusRichTextBox.Text = "Ready to convert excel file.";
            }
            else
            {
                setInactive(convertButton);
            }
        }
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

        private void resetButton_Click(object sender, EventArgs e)
        {
            this.reset();
        }
        private void setActive(object sender)
        {
            if(sender is Button)
            {
                Button b = (Button)sender;
                b.Enabled = true;
            }
        }
        private void setInactive(object sender)
        {
            if (sender is Button)
            {
                Button b = (Button)sender;
                b.Enabled = false;
            }
        }
    }
}
