using System.Text.RegularExpressions;

namespace ExtratorCTRC.Utils
{
  public static class ExtensionMethods
  {
    public static string HandlesHtmlText(this string input)
    {
      return input.Replace("&nbsp;", String.Empty);
    }

    public static string StripHTML(this string input)
    {
      return Regex.Replace(input, "<.*?>", String.Empty);
    }

    public static string GetBetween(string strSource, string strStart, string strEnd)
    {
      if (strSource.Contains(strStart) && strSource.Contains(strEnd))
      {
        int Start, End;
        Start = strSource.IndexOf(strStart, 0) + strStart.Length;
        End = strSource.IndexOf(strEnd, Start);
        return strSource.Substring(Start, End - Start);
      }

      return string.Empty;
    }

    public static string GetString(string strSource, string strStart)
    {
      return strSource.Substring(strSource.IndexOf(strStart) + strStart.Length);
    }


    public static List<string> ExtractFromBody(string body, string start, string end)
    {
      List<string> matched = new List<string>();

      int indexStart = 0;
      int indexEnd = 0;

      bool exit = false;

      try
      {
        while (!exit)
        {
          indexStart = body.IndexOf(start);

          if (indexStart != -1)
          {
            indexEnd = indexStart + body.Substring(indexStart).IndexOf(end);

            matched.Add(body.Substring(indexStart + start.Length, indexEnd - indexStart - start.Length));

            body = body.Substring(indexEnd + end.Length);
          }
          else
          {
            exit = true;
          }
        }
      }
      catch
      {
        return matched;
      }

      return matched;
    }

    public static string ReplaceString(string strSource, string wanted, string replaced)
    {
      return strSource.Replace(wanted, replaced);
    }

    public static List<string> ReturnOrderByStringNumber(this List<string> listStrings)
    {
      return listStrings.OrderBy(x => int.Parse(x.Split('-')[0])).ThenBy(x => int.Parse(x.Split('-')[1])).ToList();
    }

    public static string ReturnNameMonth(this string data)
    {
      data = data.Replace(".", string.Empty).Trim();

      int codmonth = Convert.ToInt16(data.Substring(2, 2));

      return MonthName(codmonth);
    }

    private static string MonthName(int month) =>
     month switch
     {
       1 => "Janeiro",
       2 => "Fevereiro",
       3 => "Março",
       4 => "Abril",
       5 => "Maio",
       6 => "Junho",
       7 => "Julho",
       8 => "Agosto",
       9 => "Setembro",
       10 => "Outubro",
       11 => "Novembro",
       12 => "Dezembro",
       _ => throw new Exception("Mês não encontrado!"),
     };

    public static string AjustNameMun(this string nomeMun) =>
      nomeMun switch
      {
        "BeloHorizonte" => "Belo Horizonte",
        "NovaLima" => "Nova Lima",
        _ => nomeMun
      };
  }
}
