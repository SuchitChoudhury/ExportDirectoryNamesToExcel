using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Resources;
using Microsoft.Office.Interop.Excel;

namespace ExportDirectoryNamesToExcel
{
    class Program
    {
       

        static void Main(string[] args)
        {
            var FolderPath = ConfigurationManager.AppSettings["FolderPath"];
           
            FetchDirectoryNames(FolderPath);
        }

        private static void FetchDirectoryNames(string path)
        {
            var directory = new DirectoryInfo(path);
            var ListOfSubdirectories=directory.GetDirectories();
           SortedDictionary<string, DateTime> ListOfSubDirectories = new SortedDictionary<string, DateTime>();
            foreach (var SubDirectory in ListOfSubdirectories)
            {
                ListOfSubDirectories[SubDirectory.Name] = SubDirectory.LastWriteTime;


            }
            WriteToExcel(ListOfSubDirectories);
        }

        private static void WriteToExcel(SortedDictionary<string, DateTime> listOfSubDirectories)
        {
            var MyExcel = new Application();
            var excelWorkBook = MyExcel.Workbooks.Add();
            Worksheet WorksheetData = excelWorkBook.ActiveSheet;   
            
            var ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            MyExcel.Visible = true;
            
            WorksheetData.Cells[1, 1] = "FileName";
            WorksheetData.Cells[1, 2] = "Date Modified";
            
            int rowIndex = 2;

            foreach (var SubDirectory in listOfSubDirectories)
            {                
               
                WorksheetData.Cells[rowIndex, 1] = SubDirectory.Key;
                WorksheetData.Cells[rowIndex, 2] = SubDirectory.Value;
                rowIndex++;
            }
            excelWorkBook.SaveAs($"{ExcelPath}/FileList.xlsx", Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlShared,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
           // MyExcel.GetSaveAsFilename(ExcelPath, "Excel Workbook,*.xlsx",1,"FileCount");
        }  
    }
}
