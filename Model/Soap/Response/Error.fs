namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization

[<CLIMutable>]
type Error =
    { [<XmlAttribute("codice")>]
      Code: int

      [<XmlElement("descrizione")>]
      Message: string }
