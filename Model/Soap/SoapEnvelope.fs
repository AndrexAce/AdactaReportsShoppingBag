namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap

open System.Xml.Serialization

[<XmlRoot("Envelope", Namespace = "http://www.w3.org/2003/05/soap-envelope")>]
[<CLIMutable>]
type SoapEnvelope =
    { [<XmlElement("Body")>]
      SoapBody: SoapBody }
