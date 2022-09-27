using ClosedXML.Excel;
using ExtratorCTRC.Models;
using ExtratorCTRC.Utils;
using HtmlAgilityPack;
using System.Text;

namespace ExtratorCTRC.Business
{
  public class LeituraArquivosBLL : ILeituraArquivosBLL
  {
    public string RetornaCamposHtml(IEnumerable<IFormFile>? arquivos)
    {
      try
      {
        return SeedXLS(PerformHtmlReading(arquivos));
      }
      catch (Exception ex)
      {
        throw new Exception(ex.Message);
      }
    }

    private List<CTRC> PerformHtmlReading(IEnumerable<IFormFile>? files)
    {
      HtmlDocument doc = new HtmlDocument();
      var listCtrcs = new List<CTRC>();

      if (files is not null)
        foreach (var file in files)
        {
          using (var reader = new StreamReader(file.OpenReadStream(), Encoding.UTF7))
          {
            CTRC ctrc = new CTRC();

            doc.Load(reader);

            string htmlOutput = doc.DocumentNode.OuterHtml.StripHTML().HandlesHtmlText();

            ctrc.DataEmissao = ExtensionMethods.GetBetween(htmlOutput, "Emissão.........:", "C");
            ctrc.DocTransportes = ExtensionMethods.ExtractFromBody(htmlOutput, "Doc.Transp.....:", "E");
            ctrc.Placas = ExtensionMethods.ExtractFromBody(htmlOutput, "PlacaCarreta...:", "T");

            var rowsMunicipios = ExtensionMethods.ExtractFromBody(htmlOutput, "Destinatário....:", "UF..............:");

            ctrc.Municipios = CreateListMunicipios(rowsMunicipios);
            ctrc.NumerosCTRC = ExtensionMethods.ExtractFromBody(htmlOutput, "NúmeroCTRC.....:", "S");
            ctrc.NumerosNFServ = ExtensionMethods.ExtractFromBody(htmlOutput, "NFServ..:", "Sé");

            var listNumeroCtrcServ = ctrc.NumerosCTRC.Concat(ctrc.NumerosNFServ).ToList().ReturnOrderByStringNumber();

            ctrc.TiposEmissoes = CreateListTipoEmissoes(ctrc.NumerosCTRC, listNumeroCtrcServ);

            ctrc.Etapas1DT = listNumeroCtrcServ;

            ctrc.Etapas2DT = ExtensionMethods.ExtractFromBody(htmlOutput, "EtapadoDT.....:", "I");

            ctrc.NFs01 = ExtensionMethods.ExtractFromBody(htmlOutput, "PesoValorMercadoria", "Pr");
  
            ctrc.ValoresFretesPesos = ExtensionMethods.ExtractFromBody(htmlOutput, "FretePeso......:", "M");

            listCtrcs.Add(ctrc);
          }
        }

      return listCtrcs;
    }

    private List<string> CreateListMunicipios(List<string> rowsMunicipios)
    {
      var listMun = new List<string>();

      foreach (var s in rowsMunicipios)
      {
        var municipio = ExtensionMethods.GetString(s, "Município.......:");

        listMun.Add(municipio.AjustNameMun());
      }

      return listMun;
    }

    private List<string> CreateListTipoEmissoes(List<string> numerosCTRC, List<string> listNumeroCtrcServ)
    {
      var tiposEmissoes = new List<string>();

      foreach (var ctrcOuserv in listNumeroCtrcServ)
      {
        if (numerosCTRC.Contains(ctrcOuserv))
          tiposEmissoes.Add("Número CTRC");
        else
          tiposEmissoes.Add("Número NF Serv");
      }

      return tiposEmissoes;
    }

    private string SeedXLS(List<CTRC> ctrcs)
    {
      var listCTRCGrouped = BuildListGroupedByDate(ctrcs);

      return CreateXLS(listCTRCGrouped);
    }

    private List<CTRC> BuildListGroupedByDate(List<CTRC> ctrcs)
    {
      var listCTRCGrouped = new List<CTRC>();

      foreach (var item in ctrcs)
      {
        if (listCTRCGrouped.Any(x => x.DataEmissao == item.DataEmissao))
        {
          var ctrc = listCTRCGrouped.FirstOrDefault(x => x.DataEmissao == item.DataEmissao);
          ctrc?.DocTransportes.AddRange(item.DocTransportes);
          ctrc?.Placas.AddRange(item.Placas);
          ctrc?.TiposEmissoes.AddRange(item.TiposEmissoes);
          ctrc?.Etapas1DT.AddRange(item.Etapas1DT);
          ctrc?.Etapas2DT.AddRange(item.Etapas2DT);
          ctrc?.Municipios.AddRange(item.Municipios);
          ctrc?.NFs01.AddRange(item.NFs01);
          ctrc?.ValoresFretesPesos.AddRange(item.ValoresFretesPesos);
        }
        else
        {
          listCTRCGrouped.Add(item);
        }
      }

      return listCTRCGrouped;
    }

    private string CreateXLS(List<CTRC> ctrcs)
    {
      try
      {
        var workbook = new XLWorkbook();

        foreach (var dadoCtrc in ctrcs)
        {
          var worksheet = workbook.Worksheets.Add($"{dadoCtrc.DataEmissao}");

          worksheet.Cell("A1").Value = "DT";
          worksheet.Cell("A1").Style.Font.FontName = "Courier New";
          worksheet.Cell("A1").Style.Font.FontSize = 10;
          worksheet.Cell("A1").Style.Font.Bold = true;
          worksheet.Cell("A1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromArgb(253, 233, 217);

          worksheet.Cell("B1").Value = "PLACA";
          worksheet.Cell("B1").Style.Font.FontName = "Courier New";
          worksheet.Cell("B1").Style.Font.FontSize = 10;
          worksheet.Cell("B1").Style.Font.Bold = true;
          worksheet.Cell("B1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("B1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("C1").Value = "TIPO DE EMISSAO";
          worksheet.Cell("C1").Style.Font.FontName = "Courier New";
          worksheet.Cell("C1").Style.Font.FontSize = 10;
          worksheet.Cell("C1").Style.Font.Bold = true;
          worksheet.Cell("C1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("C1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("D1").Value = "ETAPA 1 DT";
          worksheet.Cell("D1").Style.Font.FontName = "Courier New";
          worksheet.Cell("D1").Style.Font.FontSize = 10;
          worksheet.Cell("D1").Style.Font.Bold = true;
          worksheet.Cell("D1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("D1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("E1").Value = "ETAPA 2 DT";
          worksheet.Cell("E1").Style.Font.FontName = "Courier New";
          worksheet.Cell("E1").Style.Font.FontSize = 10;
          worksheet.Cell("E1").Style.Font.Bold = true;
          worksheet.Cell("E1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("E1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("F1").Value = "MUNICIPIO";
          worksheet.Cell("F1").Style.Font.FontName = "Courier New";
          worksheet.Cell("F1").Style.Font.FontSize = 10;
          worksheet.Cell("F1").Style.Font.Bold = true;
          worksheet.Cell("F1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("F1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("G1").Value = "NF1";
          worksheet.Cell("G1").Style.Font.FontName = "Courier New";
          worksheet.Cell("G1").Style.Font.FontSize = 10;
          worksheet.Cell("G1").Style.Font.Bold = true;
          worksheet.Cell("G1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("G1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          /* Possui Evolução no sistema
          worksheet.Cell("H1").Value = "NF2";
          worksheet.Cell("H1").Style.Font.FontName = "Courier New";
          worksheet.Cell("H1").Style.Font.FontSize = 10;
          worksheet.Cell("H1").Style.Font.Bold = true;
          worksheet.Cell("H1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("H1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("I1").Value = "NF3";
          worksheet.Cell("I1").Style.Font.FontName = "Courier New";
          worksheet.Cell("I1").Style.Font.FontSize = 10;
          worksheet.Cell("I1").Style.Font.Bold = true;
          worksheet.Cell("I1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("I1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("J1").Value = "NF4";
          worksheet.Cell("J1").Style.Font.FontName = "Courier New";
          worksheet.Cell("J1").Style.Font.FontSize = 10;
          worksheet.Cell("J1").Style.Font.Bold = true;
          worksheet.Cell("J1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("J1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
          */

          worksheet.Cell("H1").Value = "VALOR PESO FRETE";
          worksheet.Cell("H1").Style.Font.FontName = "Courier New";
          worksheet.Cell("H1").Style.Font.FontSize = 10;
          worksheet.Cell("H1").Style.Font.Bold = true;
          worksheet.Cell("H1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("H1").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);

          worksheet.Cell("I1").Value = "EMISSAO";
          worksheet.Cell("I1").Style.Font.FontName = "Courier New";
          worksheet.Cell("I1").Style.Font.FontSize = 10;
          worksheet.Cell("I1").Style.Font.Bold = true;
          worksheet.Cell("I1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("I1").Style.Fill.BackgroundColor = XLColor.FromArgb(221, 217, 196);

          worksheet.Cell("J1").Value = "Nº CTE";
          worksheet.Cell("J1").Style.Font.FontName = "Courier New";
          worksheet.Cell("J1").Style.Font.FontSize = 10;
          worksheet.Cell("J1").Style.Font.Bold = true;
          worksheet.Cell("J1").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
          worksheet.Cell("J1").Style.Fill.BackgroundColor = XLColor.FromArgb(221, 217, 196);

          worksheet.RangeUsed().SetAutoFilter();

          UpdateRows(dadoCtrc, worksheet);
        }

        workbook.SaveAs(@$"C:\Espelhos\Espelhos {ctrcs?.FirstOrDefault()?.DataEmissao.ReturnNameMonth()}.xlsx");

        return $"Planilha Espelhos {ctrcs?.FirstOrDefault()?.DataEmissao.ReturnNameMonth()}.xlsx criada com sucesso! Acesse o caminho C:\\Espelhos para localizar a mesma!";
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("used by another process"))
          throw new Exception("Erro ao salvar planilha: A planilha está aberta, feche a mesma e execute o processo novamente!");
        else
          throw new Exception($"Erro ao salvar planilha: {ex}");
      }
    }

    private void UpdateRows(CTRC dadoCtrc, IXLWorksheet worksheet)
    {
      int numRowDocTransporte = 2;
      int numRowPlaca = 2;
      int numRowEmissao = 2;
      int numRowEtapa1 = 2;
      int numRowEtapa2 = 2;
      int numMunicipio = 2;
      int numRowNF01 = 2;
      int numRowValorFrete = 2;

      foreach (var docTransporte in dadoCtrc.DocTransportes)
      {
        worksheet.Cell($"A{numRowDocTransporte}").Value = docTransporte;
        worksheet.Cell($"A{numRowDocTransporte}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"A{numRowDocTransporte}").Style.Font.FontSize = 10;
        worksheet.Cell($"A{numRowDocTransporte}").Style.Font.Bold = true;
        worksheet.Cell($"A{numRowDocTransporte}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"A{numRowDocTransporte}").Style.Fill.BackgroundColor = XLColor.FromArgb(253, 233, 217);
        numRowDocTransporte++;
      }

      foreach (var placa in dadoCtrc.Placas)
      {
        worksheet.Cell($"B{numRowPlaca}").Value = placa;
        worksheet.Cell($"B{numRowPlaca}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"B{numRowPlaca}").Style.Font.FontSize = 10;
        worksheet.Cell($"B{numRowPlaca}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"B{numRowPlaca}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        numRowPlaca++;
      }

      foreach (var tipoEmissao in dadoCtrc.TiposEmissoes)
      {
        worksheet.Cell($"C{numRowEmissao}").Value = tipoEmissao;
        worksheet.Cell($"C{numRowEmissao}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"C{numRowEmissao}").Style.Font.FontSize = 10;
        worksheet.Cell($"C{numRowEmissao}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"C{numRowEmissao}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        numRowEmissao++;
      }

      foreach (var etapa1 in dadoCtrc.Etapas1DT)
      {
        worksheet.Cell($"D{numRowEtapa1}").Value = etapa1;
        worksheet.Cell($"D{numRowEtapa1}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"D{numRowEtapa1}").Style.Font.FontSize = 10;
        worksheet.Cell($"D{numRowEtapa1}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"D{numRowEtapa1}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        numRowEtapa1++;
      }

      foreach (var etapa2 in dadoCtrc.Etapas2DT)
      {
        worksheet.Cell($"E{numRowEtapa2}").Value = etapa2;
        worksheet.Cell($"E{numRowEtapa2}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"E{numRowEtapa2}").Style.Font.FontSize = 10;
        worksheet.Cell($"E{numRowEtapa2}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"E{numRowEtapa2}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        worksheet.Style.NumberFormat.Format = "00#";
        numRowEtapa2++;
      }

      foreach (var municipio in dadoCtrc.Municipios)
      {
        worksheet.Cell($"F{numMunicipio}").Value = municipio;
        worksheet.Cell($"F{numMunicipio}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"F{numMunicipio}").Style.Font.FontSize = 10;
        worksheet.Cell($"F{numMunicipio}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"F{numMunicipio}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        numMunicipio++;
      }

      foreach (var nf01 in dadoCtrc.NFs01)
      {
        worksheet.Cell($"G{numRowNF01}").Value = nf01;
        worksheet.Cell($"G{numRowNF01}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"G{numRowNF01}").Style.Font.FontSize = 10;
        worksheet.Cell($"G{numRowNF01}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"G{numRowNF01}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        numRowNF01++;
      }

      foreach (var valorFretePeso in dadoCtrc.ValoresFretesPesos)
      {
        var vlrFrete = valorFretePeso.Replace(',', '.');
        worksheet.Cell($"H{numRowValorFrete}").Value = vlrFrete;
        worksheet.Cell($"H{numRowValorFrete}").Style.Font.FontName = "Courier New";
        worksheet.Cell($"H{numRowValorFrete}").Style.Font.FontSize = 10;
        worksheet.Cell($"H{numRowValorFrete}").Style.Font.FontColor = XLColor.FromArgb(0, 112, 192);
        worksheet.Cell($"H{numRowValorFrete}").Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217);
        worksheet.Cell($"H{numRowValorFrete}").Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);
        numRowValorFrete++;
      }

      worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
      worksheet.Columns(2, 20).AdjustToContents();
    }
  }
}
