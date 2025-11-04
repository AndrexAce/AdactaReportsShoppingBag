namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization

[<CLIMutable>]
type GetJobResult =
    { [<XmlAttribute("numeroJob")>]
      JobCode: string

      [<XmlAttribute("titoloDelJob")>]
      JobTitle: string

      [<XmlArray("prodotti")>]
      [<XmlArrayItem("Prodotto")>]
      Products: Product array

      [<XmlElement("errore")>]
      Error: Error }
