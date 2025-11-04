namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap

open System.Xml.Serialization
open System.Xml

[<CLIMutable>]
type SoapBody =
    { [<XmlAnyElement>]
      SoapAction: XmlNode }
