

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Novacode;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Drawing.Imaging;


namespace docManipulation
{
    class DocX_Examples
    {
        public String[] listItem = new String[20];
        public void CreateSampleDocument()
        {
            // Modify to suit your machine:
            string fileName = @"D:\DocManipulation\wordListUKFinal.docx";

            string headlineText = "Word List:";
            string paraOne = ""
                + "We the People of the United States, in Order to form a more perfect Union, "
                + "establish Justice, insure domestic Tranquility, provide for the common defence, "
                + "promote the general Welfare, and secure the Blessings of Liberty to ourselves "
                + "and our Posterity, do ordain and establish this Constitution for the United "
                + "States of America.";

            // A formatting object for our headline:
            var headLineFormat = new Formatting();
            headLineFormat.FontFamily = new System.Drawing.FontFamily("Arial Black");
            headLineFormat.Size = 18D;
            headLineFormat.Position = 12;

            // A formatting object for our normal paragraph text:
            var paraFormat = new Formatting();
            paraFormat.FontFamily = new System.Drawing.FontFamily("Calibri");
            paraFormat.Size = 10D;

            
            // Create the document in memory:
            var doc = DocX.Create(fileName);

            
            // Insert the now text obejcts;
            doc.InsertParagraph(headlineText, false, headLineFormat);
            //doc.InsertParagraph(paraOne, false, paraFormat);

            
            Read_From_Excel readExcel = new Read_From_Excel();
            readExcel.getExcelFile(@"D:\language\contentAdilBhaiUK.xlsx",2050,3);
            //readExcel.getExcelFile(@"D:\language\Test.xlsx", 100, 3);


            int rowNo = 2050 + 1;
            int columnNo = 4;
            
            Table t = doc.AddTable(rowNo, columnNo);

            var tableFormat = new Formatting();
            tableFormat.FontFamily = new System.Drawing.FontFamily("Nirmala UI");

            

            t.TableCaption = "WORD";
            //t.TableDescription


            t.Rows[0].Cells[0].Paragraphs.First().Append("Serial No.").Bold();

            t.Rows[0].Cells[1].Paragraphs.First().Append("Word").Bold();

            t.Rows[0].Cells[2].Paragraphs.First().Append("Explanation").Bold();

            t.Rows[0].Cells[3].Paragraphs.First().Append("Image").Bold();

            for (int i = 1; i < rowNo; i++)
            {
                for (int j = 0; j < columnNo; j++)
                {
                    if (j == 0)
                    {
                        int serial = i;

                        t.Rows[i].Cells[j].Paragraphs.First().Append(serial+"");
                        
                    
                    }
                    if (j == 1)
                    {
                        t.Rows[i].Cells[j].Paragraphs.First().Append(readExcel.wordList[i-1]);
                    
                    }
                    if (j == 2)
                    {
                        t.Rows[i].Cells[j].Paragraphs.First().Append(readExcel.wordListBangla[i-1]).Font(tableFormat.FontFamily);
                    
                    }
                    if (j == 3)
                    {
                        String s = readExcel.wordList[i-1];

                        Console.WriteLine("checking for::::" + s);
                        if (s.Contains(" "))
                        {
                            Console.WriteLine("word contains SPACE::::::::::::::" + s);
                            s = s.Replace(" ", "-");
                        }
                        Console.WriteLine("checking for actually::::" + s);


                        String imageFileName = imageExists(s);

                        
                        if (!imageFileName.Contains("noSuchFile"))
                        {
                            Novacode.Image image = doc.AddImage(@imageFileName);
                            Picture picture = image.CreatePicture();

                            picture.Width = 100 ;//(int)t.Rows[i].Cells[j].Width;
                            picture.Height = 100;

                            t.Rows[i].Cells[j].Paragraphs.First().AppendPicture(picture);
                        }
                        else
                        {
                            t.Rows[i].Cells[j].Paragraphs.First().Append(""); 
                        
                        }
                    }
                }
            
            }

            /*
            t.Rows[0].Cells[0].Paragraphs.First().Append("A");
            t.Rows[0].Cells[1].Paragraphs.First().Append("B");
            t.Rows[0].Cells[2].Paragraphs.First().Append("C");
            t.Rows[1].Cells[0].Paragraphs.First().Append("D");
            
            Novacode.Image image = doc.AddImage(@"D:\language\Images\17-accountant-gif.gif");
            Picture picture = image.CreatePicture();

            t.Rows[1].Cells[2].Paragraphs.First().AppendPicture(picture);

            String s = "accountant";
            String imageFileName = imageExists(s);

            Novacode.Image image1 = doc.AddImage(@imageFileName);
            Picture picture1 = image1.CreatePicture();

            t.Rows[1].Cells[1].Paragraphs.First().AppendPicture(picture1);
            */

          
            
            doc.InsertTable(t);

            
           // Paragraph p1 = doc.InsertParagraph();
           // p1.AppendPicture(picture).AppendLine("hello");

            //}

            // Save to the output directory:
            doc.Save();

            // Open in Word:
            Process.Start("WINWORD.EXE", fileName);

            
        }

        public String imageExists(String name)
        { 
            String ProcessingDirectory=@"D:\language\Images";
            DirectoryInfo di = new DirectoryInfo(ProcessingDirectory);

            
            FileInfo[] TXTFiles = di.GetFiles("*-"+name+"-*");

            String result = "";

            if (TXTFiles.Length == 0)
            {
                Console.WriteLine("no files present");
                result = "noSuchFile";
            }
            foreach (var fi in TXTFiles)
            {
                Console.WriteLine("file Exists:" + fi.Exists);
                Console.WriteLine("fileName:"+fi.FullName);
                result = fi.FullName;
            }

           // bool exist = Directory.EnumerateFiles(ProcessingDirectory, "*-"+name+"-*").Any();
           // Console.WriteLine("file Exists BOOL:" + exist);

            return result;

        }


    }
}
