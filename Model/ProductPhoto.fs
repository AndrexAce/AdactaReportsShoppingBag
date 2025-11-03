namespace AdactaInternational.AdactaReportsShoppingBag.Model

open System.Xml.Serialization

[<XmlType("FotoProdotto")>]
type ProductPhoto =
    { [<XmlElement("tipo")>]
      Type: string

      [<XmlElement("urlFoto")>]
      PhotoUrl: string }
