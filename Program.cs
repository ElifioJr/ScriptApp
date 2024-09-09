using Xceed.Words.NET;
using System;
using System.Drawing;
using Xceed.Document.NET;
using OfficeOpenXml;
using ScriptApp;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.Commercial;


        //Console.WriteLine("Informe o caminho do arquivo .xlsx");
        var excelFile = "C:\\Users\\Elifio\\Desktop\\ScriptTesteApp\\ArqTeste.xlsx";

        //Console.WriteLine("Informe o caminho do arquivo Word");
        var documentoNovo = "C:\\Users\\Elifio\\Desktop\\ScriptTesteApp\\Teste.docx";

        Console.WriteLine("Data inicio");
        var dataInicio = Console.ReadLine();

        Console.WriteLine("Data Fim");
        var dataFim = Console.ReadLine();

        var img = "img.png";

        List<DadosExcel> listaDadosExcel = new List<DadosExcel>();
        List<DadosWord> listaDadosWord = new List<DadosWord>();
        
        using (ExcelPackage package = new ExcelPackage(excelFile)) 
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            IDictionary<string, int> listaTopico = new Dictionary<string, int>();

            int linhasNulas = 0;
            for (int row = 3; row < rowCount; row++)
            {
                if (worksheet.Cells[row, 1].Value is null)
                {
                    linhasNulas++;
                    if (linhasNulas > 1)
                        break;


                    DadosWord dadosWord = new()
                    {
                        Topico = listaTopico.Aggregate((x, y) => x.Value > y.Value ? x : y).Key,
                        ListaDadosExcel = listaDadosExcel
                    };

                    listaDadosWord.Add(dadosWord);
                    listaDadosExcel.Clear();

                    continue;
                }
                linhasNulas = 0;

                // Adicionar topicos de recorrencia no dicionario;

                string step = worksheet.Cells[row, 1]?.Value?.ToString();
                string macro = worksheet.Cells[row, 2]?.Value?.ToString();
                string process = worksheet.Cells[row, 3]?.Value?.ToString();
                string action = worksheet.Cells[row, 5]?.Value?.ToString();
                string resultado = worksheet.Cells[row, 8]?.Value?.ToString();

                DadosExcel dadosExcel = new DadosExcel
                {
                    macroCenario = macro,
                    processo = process,
                    resultadoEsperado = resultado,
                    acao = action,
                    numeroStep = step
                };
                listaDadosExcel.Add(dadosExcel);

                foreach (var item in listaDadosExcel)
                {
                    if (listaTopico.ContainsKey(item.macroCenario))
                    {
                        listaTopico[item.macroCenario]++;
                    }
                    else
                    {
                        listaTopico[item.macroCenario] = 1;
                    }
                }
            }
        }

        using (var document = DocX.Create(documentoNovo))
        {
            var paragrafo = document.InsertParagraph();
            var imagem = document.AddImage(img);
            var picture = imagem.CreatePicture();

            paragrafo.IndentationBefore = 2.0f;
            paragrafo.IndentationAfter = 1.0f;

            paragrafo.AppendPicture(picture);

            paragrafo.Append("\t\t\t");
            var text = "Evidência de Testes\r\n\t\t\t\tCheckList dos Testes do Sistema\r\n";
            paragrafo.Append(text).Font("Calibri").FontSize(10);

            var espacamento = "\r\n\r\n\r\n\r\n";
            var paragrafo1 = document.InsertParagraph(espacamento);
            var paragrafo2 = document.InsertParagraph();

            var tituloDocumento = "EVIDÊNCIA DE TESTES";
            paragrafo2.Append(tituloDocumento).Font("Calibri").FontSize(28).Alignment = Xceed.Document.NET.Alignment.center;

            paragrafo2.Append(espacamento);
            paragrafo2.Append("Data Inicio: " + dataInicio + "\r\n\r\n" + "Data Fim: " + dataFim).Font("Calibri").FontSize(14);

            document.InsertParagraph().InsertPageBreakAfterSelf();
            var paragrafo3 = document.InsertParagraph("Sumario").Bold().FontSize(16);
            paragrafo3.Alignment = Xceed.Document.NET.Alignment.center;

            document.InsertParagraph().InsertPageBreakAfterSelf();

            foreach (var row in listaDadosExcel) 
            {

                var titulo = document.InsertParagraph(row.acao);
                titulo.StyleId = "Heading1";
                var table = document.AddTable(3, 1);

                table.Alignment = Xceed.Document.NET.Alignment.center;
                table.Rows[0].Cells[0].Paragraphs[0].Append("Macro Cenário: " + row.macroCenario).Font("Calibri");
                table.Rows[1].Cells[0].Paragraphs[0].Append("Processo: " + row.processo);
                table.Rows[2].Cells[0].Paragraphs[0].Append("Resultado Esperado: " + row.resultadoEsperado);
                document.InsertParagraph().InsertTableAfterSelf(table);
                document.InsertParagraph();

            }

            document.Save();
        }
    }
}
