﻿using System;
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
            DataSet testeXML = new DataSet();
            XmlDocument document = new XmlDocument();
            //document.Load(@"C:\Users\INFIS001061\OneDrive - INFIS CONSULTORIA LTDA - ME\Documentos\Roberto\XML DA DI.xml");
            document.Load(@"C:\Users\INFIS001061\OneDrive - INFIS CONSULTORIA LTDA - ME (1)\Documentos\Roberto\XML DA DI - exemplo 3.xml");
            //document.Load(@"C:\Users\INFIS001061\Documents\Roberto\XML DA DI EXEMPLO 5.xml");


            //var xdocResposta = XDocument.Parse(document.OuterXml);


            testeXML.ReadXml(new XmlTextReader(new StringReader(document.OuterXml)));
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
    }
}