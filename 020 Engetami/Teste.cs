using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Toto(string[] args)
    {
        string diretorio = Directory.GetCurrentDirectory() + @"\Relatorios020";
        Console.WriteLine("Diretório atual: " + diretorio);
        // Nome do arquivo de destino
        string nomeArquivoDestino = "dados_combinados.xlsx";
        try
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            excel.Workbooks.Add("");
            foreach (string file in Directory.GetFiles(diretorio, "*.xlsx"))
            {
                Excel.Workbook wb = excel.Workbooks.Open(file);
                Excel.Worksheet ws = wb.Worksheets[1];
                ws.Copy(After: excel.Workbooks[1].Worksheets[excel.Workbooks[1].Worksheets.Count]);
            }
            excel.Workbooks[1].SaveAs(diretorio + @"\" + nomeArquivoDestino);
            excel.Workbooks[1].Close();
            excel.Quit();


        }
        catch (Exception e)
        {
            Console.WriteLine("Erro: " + e.Message);
        }
    }
}