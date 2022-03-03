using Grpc.Core;
using OfficeOpenXml;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
//using OfficeOpenXml;
namespace testeXML
{
    class Program
    {
        
        public static decimal vlrAduaneiro = 0, vlrReceita =0;
        public static string ref_interreight, conhecCarga, conhecMaster;
        public static CultureInfo culture = CultureInfo.CreateSpecificCulture("en-CA");
        public static void Main(string[] args)
        {
           

           
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DataSet testeXML = new DataSet();
            XmlDocument document = new XmlDocument();

            document.Load(@"C:\Users\infis001061\Documents\Roberto\Projetos\TesteArquivoXML\testeXML\XML\2119761450.xml");
            var package = new ExcelPackage(new FileInfo(@"C:\Users\infis001061\Documents\Roberto\Projetos\TesteArquivoXML\testeXML\Excel\De-Para_Loreal_Modelo.xlsx"));
            testeXML.ReadXml(new XmlTextReader(new StringReader(document.OuterXml)));

            ExcelWorksheet sheetInterfreight = package.Workbook.Worksheets["Planilha Interfreight (2)"];
            ExcelWorksheet sheetCusto = package.Workbook.Worksheets["Custo (2)"];
            DataRow tbDeclacao = testeXML.Tables["declaracaoImportacao"].Rows[0];
            DataTable DtAdicao = testeXML.Tables["adicao"];
            DataTable dtPagamento = testeXML.Tables["pagamento"];
            DataTable dtMercadoria = testeXML.Tables["mercadoria"];
            int numItem = testeXML.Tables["mercadoria"].Rows.Count;
            pegarValorAduaneiro(dtPagamento, numItem);
            editSheetInterfreight(sheetInterfreight, tbDeclacao, DtAdicao, dtPagamento, dtMercadoria, numItem);
            editSheetCusto(sheetCusto, tbDeclacao, DtAdicao, dtMercadoria,numItem);
            package.SaveAs(new FileInfo(@"C:\Users\infis001061\Documents\Roberto\Projetos\TesteArquivoXML\testeXML\Excel\sampleEpplus.xlsx"));
            //document.Load(@"C:\Users\INFIS001061\Documents\Roberto\XML DA DI EXEMPLO 5.xml");


            //var xdocResposta = XDocument.Parse(document.OuterXml);



            //            DataRow row = testeXML.Tables[0].Rows[0];
            //            DataTable tabelaAdicao = testeXML.Tables["adicao"];
            //            //DataTable tabelaMercadoria = testeXML.Tables[4];



            //            var testeDoteste = row["icms"];
            //            string frase = "Gosto\n muito de\n C#";
            //            Console.WriteLine(frase);
            //            // remove as quebras de linhas
            //            frase = frase.Replace("\n", "");
            //            frase = frase.Replace("\r", "");
            //            string ttt = "01600";
            //            decimal teste = decimal.Parse(ttt) / 10000;
            //            Console.WriteLine(frase);
            //            DataTable tabelaMercadoria = testeXML.Tables[4];
            //            foreach (DataRow rowMercadoria in tabelaMercadoria.Rows)
            //            {
            //                Console.WriteLine("quantidade:" + int.Parse(rowMercadoria["quantidade"].ToString()) / 100000);
            //                Console.WriteLine("Valor Unitario:" + decimal.Parse(rowMercadoria["valorUnitario"].ToString()) / 10000000);
            //                //Console.WriteLine("Decimal :" + decimal.Parse(int.Parse(rowMercadoria["quantidade"].ToString()).ToString()));
            //                //Console.WriteLine( rowMercadoria["quantidade"].ToString().TrimStart('0').Count()); 
            //            }


            //            var td =
            //                 (from tbAdicao in tabelaAdicao.AsEnumerable()
            //                  join tbMercadoria
            //in tabelaMercadoria.AsEnumerable() on tbAdicao["adicao_Id"] equals tbMercadoria["adicao_Id"]
            //                  select (tbAdicao, tbMercadoria)).ToList();
            //            //.ToList();
            //            //DataRow rr = row.Rows[0];cargaPesoBruto

            //            var t = td[0].tbMercadoria;

            Console.WriteLine("Hello World!");
        }

        private static void editSheetCusto(ExcelWorksheet sheetCusto, DataRow tbDeclacao, DataTable dtAdicao, DataTable dtMercadoria, int numItem)
        {
            var innerJoinAdicaoItem =
                              (from tbAdicao in dtAdicao.AsEnumerable()
                               join tbMercadoria
            in dtMercadoria.AsEnumerable() on tbAdicao["adicao_Id"] equals tbMercadoria["adicao_Id"]
                               select (tbAdicao, tbMercadoria)).ToList();
            string dolar = pegarDolar(tbDeclacao["informacaoComplementar"].ToString());
            if ( numItem < 22)
            {
                sheetCusto.Cells["C15"].Value = dolar;
                sheetCusto.Cells["C16"].Value = dolar;
                sheetCusto.Cells["C50"].Value = vlrReceita;
                sheetCusto.Cells["C20"].Value = (decimal.Parse(int.Parse(tbDeclacao["cargaPesoBruto"].ToString()).ToString("F2", culture)) / 10000000).ToString().Replace(",", ".");
                sheetCusto.Cells["C21"].Value = (decimal.Parse(int.Parse(tbDeclacao["cargaPesoLiquido"].ToString()).ToString("F2", culture)) / 10000000).ToString().Replace(",", ".");
                int i = 11;
                foreach (DataRow rowAdicao in dtAdicao.Rows)
                {

                    foreach (var rowMercadoria in innerJoinAdicaoItem.Where(x => x.tbAdicao["adicao_Id"].ToString() == rowAdicao["adicao_Id"].ToString()).Select(x => x.tbMercadoria))
                    {
                        //(Valor Aduaneiro + II)*< ipiAliquotaAdValorem > dadosMercadoriaCodigoNcm
                        string[] codigoDescricao = pegarCodigoDescriao(rowMercadoria["descricaoMercadoria"].ToString());
                        sheetCusto.Cells["D" + i.ToString()].Value = int.Parse(rowAdicao["dadosMercadoriaCodigoNcm"].ToString());
                        sheetCusto.Cells["E" + i.ToString()].Value = codigoDescricao[0];
                        sheetCusto.Cells["F" + i.ToString()].Value = codigoDescricao[1];
                        sheetCusto.Cells["G" + i.ToString()].Value = decimal.Parse(int.Parse(rowMercadoria["quantidade"].ToString()).ToString()) / 100000;
                        sheetCusto.Cells["I" + i.ToString()].Value = decimal.Parse(rowMercadoria["valorUnitario"].ToString()) / 10000000;
                        //sheetCusto.Cells["G" + i.ToString()].Value = int.Parse(rowMercadoria["quantidade"].ToString()) / 100000;
                        i++;
                    }
                    
                }
            }
            else
            {

            };
        }

        private static string pegarDolar(string InfoComplementar)
        {
            var list = InfoComplementar.Split("EURO/COM.EUROPEIA");
            var listDolar = list[1].Split("\n");
            return listDolar[0];
        }

        private static string[] pegarCodigoDescriao(string rowMercadoria)
        {
            var listDescricao = rowMercadoria.Split("-");

            return listDescricao;
        }

        private static void editSheetInterfreight(ExcelWorksheet sheetInterfreight, DataRow rowDeclacao, DataTable dtAdicao, DataTable dtPagamento, DataTable dtMercadoria, int numItem)
        {
            pegarInfoComplementares(rowDeclacao["informacaoComplementar"].ToString());
       
            var numcelula = numItem + 2;

            sheetInterfreight.Cells["B3:B" + numcelula.ToString()].Value = rowDeclacao["importadorNumero"].ToString();
            sheetInterfreight.Cells["C3:C" + numcelula.ToString()].Value = rowDeclacao["dataRegistro"].ToString();
            sheetInterfreight.Cells["D3:D" + numcelula.ToString()].Value = ref_interreight;
            sheetInterfreight.Cells["G3:G" + numcelula.ToString()].Value = conhecMaster;
            sheetInterfreight.Cells["h3:H" + numcelula.ToString()].Value = conhecCarga;
            
            sheetInterfreight.Cells["E3:E" + numcelula.ToString()].Value = rowDeclacao["numeroDI"].ToString();
            sheetInterfreight.Cells["O3:O" + numcelula.ToString()].Value = vlrAduaneiro.ToString();
            sheetInterfreight.Cells["K3:K" + numcelula.ToString()].Value = rowDeclacao["viaTransporteNome"].ToString();
            sheetInterfreight.Cells["L3:L" + numcelula.ToString()].Value = rowDeclacao["tipoDeclaracaoNome"].ToString();
            sheetInterfreight.Cells["M3:M" + numcelula.ToString()].Value = rowDeclacao["importadorNome"].ToString();
            
            sheetInterfreight.Cells["T3:T" + numcelula.ToString()].Value = "18";
            var innerJoinAdicaoItem =
                              (from tbAdicao in dtAdicao.AsEnumerable()
                               join tbMercadoria
            in dtMercadoria.AsEnumerable() on tbAdicao["adicao_Id"] equals tbMercadoria["adicao_Id"]
                               select (tbAdicao, tbMercadoria)).ToList();
            int i = 3;
            int acum = 3;
            foreach (DataRow rowMercadoria in dtMercadoria.Rows)
            {

                foreach (var rowAdicao in innerJoinAdicaoItem.Where(x => x.tbAdicao["adicao_Id"].ToString() == rowMercadoria["adicao_Id"].ToString()).Select(x => x.tbAdicao))
                {
                    //(Valor Aduaneiro + II)*< ipiAliquotaAdValorem >
                    sheetInterfreight.Cells["P" + i.ToString()].Value = (decimal.Parse(int.Parse(rowAdicao["iiAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000).ToString();
                    sheetInterfreight.Cells["Q" + i.ToString()].Value = (decimal.Parse(int.Parse(rowAdicao["ipiAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000).ToString();
                    sheetInterfreight.Cells["R" + i.ToString()].Value = (decimal.Parse(int.Parse(rowAdicao["pisPasepAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000).ToString();
                    sheetInterfreight.Cells["S" + i.ToString()].Value = (decimal.Parse(int.Parse(rowAdicao["cofinsAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000).ToString();
                    sheetInterfreight.Cells["X" + i.ToString()].Value = vlrAduaneiro + (decimal.Parse(int.Parse(rowAdicao["iiAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000) * (decimal.Parse(int.Parse(rowAdicao["pisPasepAliquotaAdValorem"].ToString()).ToString("F2", culture)) / 10000);
                    
                }

                preencherValorAduaneiro(rowMercadoria, i, sheetInterfreight);
                i++;
            }
                        
            }

        private static void preencherValorAduaneiro(DataRow rowMercadoria, int i,ExcelWorksheet sheetInterfreight)
        {
            int quant = int.Parse(rowMercadoria["quantidade"].ToString()) / 100000;
            decimal vlrUnit = decimal.Parse(int.Parse(rowMercadoria["valorUnitario"].ToString()).ToString()) / 10000000;
            sheetInterfreight.Cells["N" + i.ToString()].Value = (quant * vlrUnit).ToString();
            sheetInterfreight.Cells["AG" + i.ToString()].Value = (quant * vlrUnit).ToString();
            Console.WriteLine(quant * vlrUnit);
        }

        private static void pegarValorAduaneiro(DataTable dtPagamento, int numitem)
        {
            foreach (DataRow linha in dtPagamento.Rows)
            {
                switch (linha["codigoReceita"])
                {
                    case "7811":
                        vlrReceita = decimal.Parse(int.Parse(linha["valorReceita"].ToString()).ToString()) / 100;
                        vlrAduaneiro = ((int.Parse(linha["valorReceita"].ToString()) / 100) / numitem); //valor unitario * quantidade 
                        break;

                };

            }
        }

        private static void pegarInfoComplementares(string rowDeclacao)
        {
            //ref_interreight
            string split = "\n";
            var list_part1 = rowDeclacao.Split("REF. INTERFREIGHT.:");
            var list_ref = list_part1[1].Split(split);
            var list_part2 = list_part1[1].Split("CONHECIMENTO DE CARGA.:");
            var lis_conhecCarga = list_part2[1].Split(split);
            var list_part3 = list_part2[1].Split("CONHECIMENTO MASTER:");
            var list_conhecMaster = list_part3[1].Split(split);
            ref_interreight = list_ref[0];
            conhecCarga = lis_conhecCarga[0];
            conhecMaster = list_conhecMaster[0];

        }
    }

}

