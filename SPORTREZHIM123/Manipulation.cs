using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPORTREZHIM123
{
    internal class Manipulation
    {
        public string[,] data()
        {

            string excelFilePath = @"C:\Users\admin\Downloads\1.xlsx";

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                int rowCount = end.Row - start.Row + 1;
                int columnCount = end.Column - start.Column + 1;

                string[,] dataArray = new string[rowCount, columnCount];

                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        dataArray[row - start.Row, col - start.Column] = Convert.ToString(worksheet.Cells[row, col].Value);
                    }
                }
                return dataArray;
            }
        public double weekend(string[,] dataArray)
        {
            double sum = 0;
            for (int i = 1; i < 32; i++)
            {
                if (dataArray[i, 1] == "Суббота" || dataArray[i, 1] == "Воскресенье") {
                    sum += Convert.ToDouble(dataArray[i, 3]);
            }
            }
            return sum;
        }

        }

    }
}
