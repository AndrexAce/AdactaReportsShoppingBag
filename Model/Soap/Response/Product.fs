namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel.DataAnnotations
open AdactaInternational.AdactaReportsShoppingBag.Model
open Newtonsoft.Json
open CommunityToolkit.Mvvm.ComponentModel

type Product() =
    inherit ObservableObject()
    
    let mutable _code = ""
    let mutable _name = ""
    let mutable _category = ""
    let mutable _manufacturer = ""
    let mutable _format = ""
    let mutable _ean = ""
    let mutable _batch = ""
    let mutable _productionDate = ""
    let mutable _expiryDate = ""
    let mutable _productionFacility = ""
    let mutable _ingredients = ""
    let mutable _productPhotos: ProductPhoto array = [||]
    let mutable _classification = ProductClassification.Unknown
    
    [<XmlAttribute("sigla")>]
    [<Required>]
    [<RegularExpression(@"^[A-Z0-9]{3}$")>]
    member this.Code
        with get() = _code
        and set(value) =
            if this.SetProperty(&_code, value) then
                this.OnPropertyChanged("DisplayName")

    [<XmlAttribute("nome")>]
    [<Required>]
    [<StringLength(100, MinimumLength = 1)>]
    member this.Name
        with get() = _name
        and set(value) =
            if this.SetProperty(&_name, value) then
                this.OnPropertyChanged("DisplayName")
                this.OnPropertyChanged("ProductName")
                this.OnPropertyChanged("Brand")
                this.OnPropertyChanged("SubBrand")

    [<XmlAttribute("categoria")>]
    [<Required>]
    member this.Category
        with get() = _category
        and set(value) = this.SetProperty(&_category, value) |> ignore

    [<XmlAttribute("produttore")>]
    [<Required>]
    member this.Manufacturer
        with get() = _manufacturer
        and set(value) = this.SetProperty(&_manufacturer, value) |> ignore

    [<XmlAttribute("formato")>]
    [<Required>]
    member this.Format
        with get() = _format
        and set(value) = this.SetProperty(&_format, value) |> ignore

    [<XmlAttribute("EAN")>]
    [<Required>]
    member this.EAN
        with get() = _ean
        and set(value) = this.SetProperty(&_ean, value) |> ignore

    [<XmlAttribute("lotto")>]
    [<Required>]
    member this.Batch
        with get() = _batch
        and set(value) = this.SetProperty(&_batch, value) |> ignore

    [<XmlAttribute("dataDiProduzione")>]
    [<Required>]
    member this.ProductionDate
        with get() = _productionDate
        and set(value) = this.SetProperty(&_productionDate, value) |> ignore

    [<XmlAttribute("dataDiScadenza")>]
    [<Required>]
    member this.ExpiryDate
        with get() = _expiryDate
        and set(value) = this.SetProperty(&_expiryDate, value) |> ignore

    [<XmlAttribute("stabilimentoDiProduzione")>]
    [<Required>]
    member this.ProductionFacility
        with get() = _productionFacility
        and set(value) = this.SetProperty(&_productionFacility, value) |> ignore

    [<XmlElement("ingredienti")>]
    [<Required>]
    member this.Ingredients
        with get() = _ingredients
        and set(value) = this.SetProperty(&_ingredients, value) |> ignore

    [<XmlArray("foto")>]
    [<XmlArrayItem("FotoProdotto")>]
    [<Required>]
    [<Length(1, 5)>]
    member this.ProductPhotos
        with get() = _productPhotos
        and set(value) = this.SetProperty(&_productPhotos, value) |> ignore

    [<XmlIgnore>]
    member this.Classification 
        with get() = _classification
        and set(value) = 
            if this.SetProperty(&_classification, value) then
                this.OnPropertyChanged("NavigationMenuItemIcon")

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