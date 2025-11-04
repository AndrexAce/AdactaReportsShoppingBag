namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Request

open System.Xml.Serialization

[<CLIMutable>]
[<XmlRoot(Namespace = "http://tempuri.org/")>]
type GetJob =
    { [<XmlElement("numeroJob")>]
      JobCode: string

      [<XmlElement("usr")>]
      Username: string

      [<XmlElement("pss")>]
      Password: string }
