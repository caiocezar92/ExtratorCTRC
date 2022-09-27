namespace ExtratorCTRC.Business
{
  public interface ILeituraArquivosBLL
  {
    string RetornaCamposHtml(IEnumerable<IFormFile>? arquivos);
  }
}
