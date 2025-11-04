namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization

[<CLIMutable>]
[<XmlRoot(Namespace = "http://tempuri.org/")>]
type GetJobResponse =
    { [<XmlElement("GetJobResult")>]
      GetJobResult: GetJobResult }
