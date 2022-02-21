using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
//using OfficeOpenXml;
namespace testeXML
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DataSet testeXML = new DataSet();
            XmlDocument document = new XmlDocument();
            //document.Load(@"C:\Users\INFIS001061\OneDrive - INFIS CONSULTORIA LTDA - ME\Documentos\Roberto\XML DA DI.xml");
            document.Load(@"C:\Users\infis001061\Documents\Roberto\Projetos\TesteArquivoXML\testeXML\XML\2119761450.xml");
            var package = new ExcelPackage(new FileInfo(@"C:\Users\infis001061\Documents\Roberto\Projetos\TesteArquivoXML\testeXML\Excel\De-Para_Loreal_Modelo.xlsx"));
            testeXML.ReadXml(new XmlTextReader(new StringReader(document.OuterXml)));

            ExcelWorksheet sheetInterfreight = package.Workbook.Worksheets[0];
            DataRow tbDeclacao = testeXML.Tables["declaracaoImportacao"].Rows[0];

            editSheetInterfreight(sheetInterfreight, tbDeclacao);
             //document.Load(@"C:\Users\INFIS001061\Documents\Roberto\XML DA DI EXEMPLO 5.xml");


             //var xdocResposta = XDocument.Parse(document.OuterXml);



             DataRow row = testeXML.Tables[0].Rows[0];
            DataTable tabelaAdicao = testeXML.Tables["adicao"];
            //DataTable tabelaMercadoria = testeXML.Tables[4];



            var testeDoteste = row["icms"];
            string frase = "Gosto\n muito de\n C#";
            Console.WriteLine(frase);
            // remove as quebras de linhas
            frase = frase.Replace("\n", "");
            frase = frase.Replace("\r", "");
            string ttt = "01600";
            decimal teste = decimal.Parse(ttt) / 10000;
            Console.WriteLine(frase);
            DataTable tabelaMercadoria = testeXML.Tables[4];
            foreach (DataRow rowMercadoria in tabelaMercadoria.Rows)
            {
                Console.WriteLine("quantidade:" + int.Parse(rowMercadoria["quantidade"].ToString()) / 100000);
                Console.WriteLine("Valor Unitario:" + decimal.Parse(rowMercadoria["valorUnitario"].ToString()) / 10000000);
                //Console.WriteLine("Decimal :" + decimal.Parse(int.Parse(rowMercadoria["quantidade"].ToString()).ToString()));
                //Console.WriteLine( rowMercadoria["quantidade"].ToString().TrimStart('0').Count()); 
            }
            

            var  td =
                 (from tbAdicao in tabelaAdicao.AsEnumerable() join tbMercadoria
                  in tabelaMercadoria.AsEnumerable() on tbAdicao["adicao_Id"] equals tbMercadoria["adicao_Id"]
                  select (tbAdicao, tbMercadoria)).ToList();
            //.ToList();
            //DataRow rr = row.Rows[0];

            var t = td[0].tbMercadoria;

            Console.WriteLine("Hello World!");
        }

        private static void editSheetInterfreight(ExcelWorksheet sheetInterfreight, DataRow tbDeclacao)
        {
            
            throw new NotImplementedException();
        }
    }
}
