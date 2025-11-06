namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel.DataAnnotations

[<CLIMutable>]
type Product =
    { [<XmlAttribute("sigla")>]
      [<Required>]
      [<RegularExpression(@"^[A-Z0-9]{3}$")>]
      Code: string

      [<XmlAttribute("nome")>]
      [<Required>]
      [<StringLength(100, MinimumLength = 1)>]
      Name: string

      [<XmlAttribute("categoria")>]
      [<Required>]
      Category: string

      [<XmlAttribute("produttore")>]
      [<Required>]
      Manufacturer: string

      [<XmlAttribute("formato")>]
      [<Required>]
      Format: string

      [<XmlAttribute("EAN")>]
      [<Required>]
      EAN: string

      [<XmlAttribute("lotto")>]
      [<Required>]
      Batch: string

      [<XmlAttribute("dataDiProduzione")>]
      [<Required>]
      ProductionDate: string

      [<XmlAttribute("dataDiScadenza")>]
      [<Required>]
      ExpiryDate: string

      [<XmlAttribute("stabilimentoDiProduzione")>]
      [<Required>]
      ProductionFacility: string

      [<XmlElement("ingredienti")>]
      [<Required>]
      Ingredients: string

      [<XmlArray("foto")>]
      [<XmlArrayItem("FotoProdotto")>]
      [<Required>]
      [<Length(1, 5)>]
      ProductPhotos: ProductPhoto array }
