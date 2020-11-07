using System;
using System.Collections.Generic;
using System.Linq;
using LinqToExcel;

namespace DriveProjGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToExcelFile = "" + @"D:\Code\Blog Projects\BlogSandbox\ArtistAlbums.xlsx";
            string sheetName = "Sheet1";

            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            var artistAlbums = from a in excelFile.Worksheet(sheetName) select a;

            foreach (var a in artistAlbums)
            {
                string artistInfo = "Artist Name: {0}; Album: {1}";
                Console.WriteLine(string.Format(artistInfo, a["Name"], a["Title"]));
            }
            int i = 0;
        }
    }
}