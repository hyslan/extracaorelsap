using System;
using System.Windows.Forms;
using System.Linq;
using System.Threading.Tasks;
using SAPFEWSELib;
using SapSessao;
using SQLConnection;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Diagnostics;
using Microsoft.VisualBasic.FileIO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using TransformaOrdemtoCSV;
using System.Collections.Generic;

namespace _020_Engetami
{
    public class QuerySql
    {
        public readonly string QueryString;

        public QuerySql(string query)
        {
            this.QueryString = query;
        }

        public void ExtracaoOrdem(string query)
        {
            string filepath = Directory.GetCurrentDirectory() + "\\pendente_wfm.csv";

            // Verifica se o arquivo existe e o exclui
            if (File.Exists(filepath))
            {
                File.Delete(filepath);
            }
            try
            {
                SqlConnectionBdMlg connectionObj = new SqlConnectionBdMlg();
                SqlConnection connect = connectionObj.BD_MLG_Query();
                SqlCommand command = new SqlCommand(QueryString, connect);
                connect.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(dt);
                using (StreamWriter writer = new StreamWriter(filepath))
                {
                    // Escreve os dados no arquivo CSV
                    foreach (DataRow row in dt.Rows)
                    {
                        string[] fields = row.ItemArray.Select(field => field.ToString()).ToArray();
                        string line = string.Join(",", fields);
                        writer.WriteLine(line.TrimEnd(','));
                    }
                    writer.Dispose();
                }
                Console.WriteLine("Dados exportados para o arquivo CSV com sucesso.\n");
                // Utiliza o TextFieldParser para ler o arquivo CSV com as colunas delimitadas por vírgula
                TextFieldParser parser = new TextFieldParser(filepath);
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                }
                connect.Close();
                command.Dispose();
                adapter.Dispose();
                parser.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ExtracaoOrdem: {ex.Message}");
            }

        }
    }
    public class ConsultaSap
    {
        public void ExtracaoRelatorioSap(string arquivo, string un, string contratada)
        {
            SapSessao1 sessaoObj = new SapSessao1();
            GuiSession session = sessaoObj.Sessao();
            GuiFrameWindow frame = (GuiFrameWindow)session.FindById("wnd[0]");
            session.StartTransaction("ZSBPM020");
            GuiButton selecao_multipla = (GuiButton)session.FindById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH");
            selecao_multipla.Press();
            // Botão Import CSV.
            frame.SendVKey(23);
            GuiTextField csv_path = (GuiTextField)session.FindById("wnd[2]/usr/ctxtDY_PATH");
            csv_path.Text = Directory.GetCurrentDirectory();
            GuiTextField csv_file = (GuiTextField)session.FindById("wnd[2]/usr/ctxtDY_FILENAME");
            csv_file.Text = "pendente_wfm.csv";
            GuiButton btn_import = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]");
            btn_import.Press();
            Console.WriteLine("Importação concluída do CSV.");
            frame.SendVKey(8);
            GuiTextField contrato = (GuiTextField)session.FindById("wnd[0]/usr/txtS_CONTR-LOW");
            contrato.Text = contratada;
            GuiTextField unidade_administrativa = (GuiTextField)session.FindById("wnd[0]/usr/txtS_UN_ADM-LOW");
            unidade_administrativa.Text = un;
            GuiTextField limite_ocorrencias = (GuiTextField)session.FindById("wnd[0]/usr/txtP_MAX");
            limite_ocorrencias.Text = "50000";
            GuiTextField layout = (GuiTextField)session.FindById("wnd[0]/usr/ctxtP_LAYOUT");
            layout.Text = "/MLG_MEDICAO";
            frame.SendVKey(0);
            Console.WriteLine("Efetuando consulta...");
            frame.SendVKey(8);
            try
            {
                GuiGridView guiGrid = (GuiGridView)session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell");
                // Seleçao das Colunas Status.
                guiGrid.SelectColumn("STTXT");
                guiGrid.SelectColumn("USTXT");
                GuiButton btn_filtro = (GuiButton)session.FindById("wnd[0]/tbar[1]/btn[29]");
                btn_filtro.Press();
                GuiTextField filtro_status_sistema = (GuiTextField)session.FindById(
                    "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW");
                filtro_status_sistema.Text = "LIB";
                GuiTextField filtro_status_usuario = (GuiTextField)session.FindById(
                    "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW");
                filtro_status_usuario.Text = "EXEC";
                frame.SendVKey(0);
                guiGrid.ContextMenu();
                guiGrid.SelectContextMenuItem("&XXL");
                GuiButton btn_planilha = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[0]");
                btn_planilha.Press();
                GuiTextField planilha_path = (GuiTextField)session.FindById("wnd[1]/usr/ctxtDY_PATH");
                planilha_path.Text = Directory.GetCurrentDirectory()+"\\Relatorios020\\";
                GuiTextField planilha_file = (GuiTextField)session.FindById("wnd[1]/usr/ctxtDY_FILENAME");
                planilha_file.Text = arquivo;
                Console.WriteLine("Salvando planilha 020.");
                frame.SendVKey(11);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Erro em Extração ExtracaoRelatorioSap: {e.Message}");
            }
        }

        public void ExtracaoRelatorio017Sap(string arquivo)
        {
            Console.WriteLine("Iniciando Processo do Relatório 017 no SAP.\n");
            SapSessao1 sessaoObj = new SapSessao1();
            GuiSession session = sessaoObj.Sessao();
            GuiFrameWindow frame = (GuiFrameWindow)session.FindById("wnd[0]");
            session.StartTransaction("ZSBPM017");
            GuiCheckBox box_encerrado = (GuiCheckBox)session.FindById("wnd[0]/usr/chkSP_MAB");
            box_encerrado.Selected = true;
            GuiButton selecao_multipla = (GuiButton)session.FindById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH");
            selecao_multipla.Press();
            // Botão Import CSV.
            frame.SendVKey(23);
            GuiTextField csv_path = (GuiTextField)session.FindById("wnd[2]/usr/ctxtDY_PATH");
            csv_path.Text = arquivo;
            GuiTextField csv_file = (GuiTextField)session.FindById("wnd[2]/usr/ctxtDY_FILENAME");
            csv_file.Text = "017.csv";
            GuiButton btn_import = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]");
            btn_import.Press();
            Console.WriteLine("Importação concluida do CSV.");
            frame.SendVKey(8);
            GuiTextField limite_ocorrencias = (GuiTextField)session.FindById("wnd[0]/usr/txtP_MAX");
            limite_ocorrencias.Text = "50000";
            GuiTextField layout = (GuiTextField)session.FindById("wnd[0]/usr/ctxtP_LAYOUT");
            layout.Text = "/MLG_JADER";
            frame.SendVKey(0);
            Console.WriteLine("Efetuando consulta 017...");
            frame.SendVKey(8);
            GuiGridView guiGrid = (GuiGridView)session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell");
            guiGrid.ContextMenu();
            guiGrid.SelectContextMenuItem("&XXL");
            GuiButton btn_planilha = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[0]");
            btn_planilha.Press();
            GuiTextField planilha_path = (GuiTextField)session.FindById("wnd[1]/usr/ctxtDY_PATH");
            planilha_path.Text = "C:\\Users\\irgpapais\\Documents\\Valoracao\\";
            MessageBox.Show("Relatório 017 pronto. Gerar arquivo na transação ZSBPM017.", "Aviso");
            //GuiTextField planilha_file = (GuiTextField)session.FindById("wnd[1]/usr/ctxtDY_FILENAME");
            //planilha_file.Text = "017.xlsx";
            //Console.WriteLine("Salvando planilha 017.");
            //frame.SendVKey(11);
        }
    }



    internal class Program
    {
        static void Main(string[] args)
        {
            // Número do Contrato
            string novaspMLG = "4600041302";
            string gbItaquera = "4600042888";
            string NorteSulMLG = "4600043760";
            string NorteSulMLN = "4600045267";
            string NorteSulMLN2 = "4600046036";
            string NorteSulMLQ = "4600043654";
            string ZCMLN = "4600042975";
            string RecapeMLN = "4600044787";
            string RecapeMLQ = "4600044777";
            string RecapeMLG = "4600044782";

            // UGR
            string mlg = "344";
            string mlq = "340";
            string mln = "348";

            // MLG
            string[] novaspMLGQueries = GenerateQueries("NOVASP_MLG", 8);
            string[] NorteSulMLGQueries = GenerateQueries("NORTESUL_MLG", 1);
            string[] RecapeMLGQueries = GenerateQueries("RECAPE_MLG", 1);

            // MLQ
            string[] gbItaqueraQueries = GenerateQueries("GB_ITAQUERA", 6);
            string[] NorteSulMLQQueries = GenerateQueries("NORTESUL_MLQ", 1);
            string[] RecapeMLQQueries = GenerateQueries("RECAPE_MLQ", 1);

            // MLN
            string[] ZCMLNQueries = GenerateQueries("ZC_MLN", 9);
            string[] NorteSulMLNQueries = GenerateQueries("NORTESUL_MLN", 1);
            string[] NorteSulMLN2Queries = GenerateQueries("NORTESUL_MLN2", 1);
            string[] RecapeMLNQueries = GenerateQueries("RECAPE_MLN", 2);

            QuerySql[] novaspSqlMLGQueries = InitializeQueries(novaspMLGQueries);
            QuerySql[] RecapeSqlMLGQueries = InitializeQueries(RecapeMLGQueries);
            QuerySql[] NorteSulSqlMLGQueries = InitializeQueries(NorteSulMLGQueries);

            QuerySql[] gbItaqueraSqlQueries = InitializeQueries(gbItaqueraQueries);
            QuerySql[] NorteSulSqlMLQQueries = InitializeQueries(NorteSulMLQQueries);
            QuerySql[] RecapeSqlMLQQueries = InitializeQueries(RecapeMLQQueries);

            QuerySql[] ZCSqlMLNQueries = InitializeQueries(ZCMLNQueries);
            QuerySql[] NorteSulSqlMLNQueries = InitializeQueries(NorteSulMLNQueries);
            QuerySql[] NorteSulSqlMLN2Queries = InitializeQueries(NorteSulMLN2Queries);
            QuerySql[] RecapeSqlMLNQueries = InitializeQueries(RecapeMLNQueries);

            ConsultaSap sap = new ConsultaSap();
            try
            {
                ProcessQueriesAndSap(novaspSqlMLGQueries, sap, mlg, novaspMLG, "NOVASP - MLG");
                ProcessQueriesAndSap(RecapeSqlMLGQueries, sap, mlg, RecapeMLG, "Recape - MLG");
                ProcessQueriesAndSap(NorteSulSqlMLGQueries, sap, mlg, NorteSulMLG, "NorteSUL - MLG");

                ProcessQueriesAndSap(gbItaqueraSqlQueries, sap, mlq, gbItaquera, "GB Itaquera - MLQ");
                ProcessQueriesAndSap(RecapeSqlMLQQueries, sap, mlq, RecapeMLQ, "Recape - MLQ");
                ProcessQueriesAndSap(NorteSulSqlMLQQueries, sap, mlq, NorteSulMLQ, "NorteSUl - MLQ");

                ProcessQueriesAndSap(ZCSqlMLNQueries, sap, mln, ZCMLN, "ZC - MLN");
                ProcessQueriesAndSap(RecapeSqlMLNQueries, sap, mln, RecapeMLN, "Recape - MLN");
                ProcessQueriesAndSap(NorteSulSqlMLNQueries, sap, mln, NorteSulMLN, "NorteSul - MLN");
                ProcessQueriesAndSap(NorteSulSqlMLN2Queries, sap, mln, NorteSulMLN2, "NorteSul - MLN");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro de execução: {ex.Message}");
            }
            finally
            {
                // Espera abrir a última planilha Excel
                System.Threading.Thread.Sleep(5000);
                // Encerra os processos do Excel ativos
                KillExcelProcesses();
                // Criar uma instância do Excel
                Excel.Application excelApp = new Excel.Application();

                // Tornar o aplicativo Excel invisível
                excelApp.Visible = false;

                string filepath = "C:\\Users\\irgpapais\\Documents\\Valoracao\\020.XLSX";
                Excel.Workbook wb = excelApp.Workbooks.Open(filepath);
                wb.RefreshAll();
                Console.WriteLine("Atualizando Query do Excel...");
                System.Threading.Thread.Sleep(25000);

                Transform020Tocsv csv017 = new Transform020Tocsv(wb);
                csv017.OrdensemCsv(wb);

                // Fecha e salva as alterações feitas no arquivo
                wb.Close(true);
                excelApp.Quit();

                // Libera os objetos
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Executa Garbage Collector para liberar a memória ocupada pelos objetos COM
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Encerra os processos do Excel ativos
                KillExcelProcesses();
                // ListaOS entrarár no lugar da 017.
                //string caminho017 = "C:\\Users\\irgpapais\\Documents\\Valoracao";
                //sap.ExtracaoRelatorio017Sap(caminho017);
            }

        }
        static void KillExcelProcesses()
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        static string[] GenerateQueries(string prefix, int iterations)
        {
            string[] queries = new string[iterations];

            for (int i = 0; i < iterations; i++)
            {
                int startLine = i * 25000 + 1;
                int endLine = (i + 1) * 25000;

                queries[i] = $"SELECT ORDEM FROM [BD_MLG].[LESTE_AD\\hcruz_novasp].[v_Hyslancruz_BEXEC_WFM_Ordens_{prefix}]\n" +
                              $"WHERE LINHA > {startLine} AND LINHA <= {endLine}\n ORDER BY LINHA ASC";
            }

            return queries;
        }

        static QuerySql[] InitializeQueries(string[] queries)
        {
            return queries.Select(q => new QuerySql(q)).ToArray();
        }

        static void ProcessQueriesAndSap(QuerySql[] queries, ConsultaSap sap, string unadm, string contratada, string nome)
        {
            for (int i = 0; i < queries.Length; i++)
            {
                Console.Write($"Iniciando consulta no SQL {contratada} : {nome} parte {i + 1}\n");
                queries[i].ExtracaoOrdem(queries[i].QueryString);
                Console.WriteLine($"Iniciando Processo do Relatório 020 {contratada} : {nome} no SAP parte {i + 1}\n");
                sap.ExtracaoRelatorioSap($"020 {contratada} - {nome} p{i + 1}.XLSX", unadm, contratada);
            }
        }
    }
}
