namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel.DataAnnotations

[<CLIMutable>]
type ProductPhoto =
    { [<XmlElement("tipo")>]
      [<Required>]
      Type: string

      [<XmlElement("urlFoto")>]
      [<Required>]
      PhotoUrl: string }
