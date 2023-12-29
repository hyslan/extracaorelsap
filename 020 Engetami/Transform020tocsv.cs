using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TransformaOrdemtoCSV
{
    public class Transform020Tocsv
    {
        public Excel.Workbook workbook { get; set; }

        public Transform020Tocsv(Excel.Workbook wb) 
        {
            this.workbook = wb;
        }

        public void OrdensemCsv(Excel.Workbook wb)
        {
            // Selecionar a primeira planilha
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            // Encontrar a última linha usada na planilha
            int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            // Selecionar o intervalo da primeira linha até a última linha com valor
            Excel.Range range = worksheet.Range["B2", "B" + lastRow];

            // Obter os valores do intervalo em um array bidimensional
            object[,] rangeValues = range.Value;

            string csvFilePath = @"C:\Users\irgpapais\Documents\Valoracao\\017.csv";
            // Escrever os valores do array no arquivo CSV
            using (StreamWriter sw = new StreamWriter(csvFilePath))
            {
                int rows = rangeValues.GetLength(0);
                int cols = rangeValues.GetLength(1);

                for (int row = 1; row <= rows; row++)
                {
                    for (int col = 1; col <= cols; col++)
                    {
                        string value = rangeValues[row, col]?.ToString() ?? string.Empty;
                        sw.Write(value);

                        if (col < cols)
                            sw.Write(",");
                    }

                    sw.WriteLine();
                }
            }
        }
    }
}
