using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CompareEmails
{
    class Program
    {
        static void Main(string[] args)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\temp\AppRiverO365.xlsx");

            //This is where you select the sheet. 
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //Define array to hold emails we are looking for
            List<String> sourceEmails = new List<String>();

            //Define array to hold email list we are looking in.
            List<String> targetEmails = new List<String>();

            //Override Row count as we only want to print the headers. 
            rowCount = 255;


            Console.WriteLine("Reading in source Emails");
            //Read in all the source emails. i = 2 to not read in headers, row count hard coded. 
            for (int i = 2; i <= 238; i++)
            {
                    //write the value to the console for testing. 
                    if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                    {
                        String temp = xlRange.Cells[i, 3].Value2.ToString();

                        temp = temp.ToLower();
                        sourceEmails.Add(temp);
                    }
            }

            Console.WriteLine("Reading in target Emails");
            //Read in all the source emails. i = 2 to not read in headers, row count hard coded. 
            for (int i = 2; i <= 387; i++)
            {
                //write the value to the console for testing. 
                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                {
                    String temp = xlRange.Cells[i, 5].Value2.ToString();

                    temp = temp.ToLower();
                    targetEmails.Add(temp);
                }
            }

            Console.WriteLine("Comparing emails from list 1 to list 2");
            //compare emails in the list. 
            var rowCounter = 2;
            foreach (string line in sourceEmails)
            {
                var containesEmail = targetEmails.Contains(line);

                if (containesEmail)
                {
                    xlRange.Cells[rowCounter, 4].Value2 = "Yes";
                }
                else
                {
                    xlRange.Cells[rowCounter, 4].Value2 = "No";
                }

                rowCounter++;
            }

            Console.WriteLine("Comparing emails from list 2 to list 1");
            //compare emails in the list. 
            rowCounter = 2;
            foreach (string line in targetEmails)
            {

                var containesEmail = sourceEmails.Contains(line);

                if (containesEmail)
                {
                    xlRange.Cells[rowCounter, 7].Value2 = "Yes";
                }
                else
                {
                    xlRange.Cells[rowCounter, 7].Value2 = "No";
                }

                rowCounter++;

            }

            Console.WriteLine("Cleaning up");
            Console.WriteLine("When propmted please save a new copy of the excel file.");

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            //Hold open the command window.
            //Console.ReadLine();


        }
    }
}
