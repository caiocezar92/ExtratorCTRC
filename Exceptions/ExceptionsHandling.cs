using ExtratorCTRC.Utils;

namespace ExtratorCTRC.Exceptions
{
  public static class ExceptionsHandling
  {
    public static string ReturnTreatedException(this string message)
    {
      if (message.Contains("used by another process"))
        return "Erro ao salvar planilha: A planilha está aberta, feche a mesma e execute o processo novamente!";
      else if (message.Contains("A worksheet with the same name"))
      {
        var nameWorksheet = ExtensionMethods.GetBetween(message, "name", "has");
        return $"Uma planilha com o mesmo nome {nameWorksheet} já foi adicionada. Exclua a mesma e adicione novamente os espelhos";
      }
      else
        return $"Erro ao salvar planilha: {message}";
    }
  }
}
