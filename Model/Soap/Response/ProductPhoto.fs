namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization

[<CLIMutable>]
type ProductPhoto =
    { [<XmlElement("tipo")>]
      Type: string

      [<XmlElement("urlFoto")>]
      PhotoUrl: string }
