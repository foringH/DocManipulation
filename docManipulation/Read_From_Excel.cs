using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace docManipulation
{
    class Read_From_Excel
    {
        public String[] wordList;
        public String[] wordListBangla ;
        public void getExcelFile(String fileName,int rowCount,int colCount)
        {
            wordList = new String[rowCount];
            wordListBangla = new String[rowCount];

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@fileName); // (@"D:\language\Test.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Console.WriteLine("range is:"+xlRange.ToString());

            //int rowCount = 20; //xlRange.Rows.Count;
            //int colCount = 3;//xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                        if (j == 2)
                        {
                            wordList[i - 1] = xlRange.Cells[i, j].Value2.ToString();

                        }
                        if (j == 3)
                        {
                            wordListBangla[i - 1] = xlRange.Cells[i, j].Value2.ToString();
                        }
                    }
                }
            }

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

           // printList(rowCount);

           // return wordList;

        }

        public void printList(int rowCount)
        {
            Console.WriteLine("PRINTING WORDLIST:::::::::");
            for (int i = 0; i < rowCount; i++)
            {
                Console.WriteLine("i=" + i + " " + wordList[i]+"\t"+wordListBangla[i]);
            }
        
        }

    }
}
