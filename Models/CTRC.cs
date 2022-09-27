namespace ExtratorCTRC.Models
{
  public class CTRC
  {
    public string DataEmissao { get; set; }
    public List<string> DocTransportes { get; set; }
    public List<string> Placas { get; set; }
    public List<string> Municipios { get; set; }
    public List<string> NumerosNFServ { get; set; }
    public List<string> NumerosCTRC { get; set; }
    public List<string> TiposEmissoes { get; set; }
    public List<string> Etapas1DT { get; set; }
    public List<string> Etapas2DT { get; set; }
    public List<string> NFs01 { get; set; }
    public List<string> NFs02 { get; set; }
    public List<string> NFs03 { get; set; }
    public List<string> NFs04 { get; set; }
    public List<string> ValoresFretesPesos { get; set; }
  }
}
