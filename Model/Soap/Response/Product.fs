namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel.DataAnnotations
open AdactaInternational.AdactaReportsShoppingBag.Model
open Newtonsoft.Json

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
      ProductPhotos: ProductPhoto array

      [<XmlIgnore>]
      [<DefaultValue>]
      mutable Classification: ProductClassification }

    [<XmlIgnore>]
    [<JsonIgnore>]
    member this.DisplayName =
        sprintf "%s - %s" this.Code this.ProductName

    [<XmlIgnore>]
    [<JsonIgnore>]
    member this.ProductName =
        let members = this.Name.Split([| '-' |], System.StringSplitOptions.RemoveEmptyEntries)
        if members.Length >= 1 then members[0] else ""

    [<XmlIgnore>]
    [<JsonIgnore>]
    member this.Brand =
        let members = this.Name.Split([| '-' |], System.StringSplitOptions.RemoveEmptyEntries)
        if members.Length >= 2 then members[1] else ""

    [<XmlIgnore>]
    [<JsonIgnore>]
    member this.SubBrand =
        let members = this.Name.Split([| '-' |], System.StringSplitOptions.RemoveEmptyEntries)
        if members.Length >= 3 then members[2] else ""

    [<XmlIgnore>]
    [<JsonIgnore>]
    member this.NavigationMenuItemIcon =
        match this.Classification with
        | ProductClassification.Food -> "\ued56"
        | ProductClassification.NonFood -> "\ue80f"
        | _ -> "\ue897"
