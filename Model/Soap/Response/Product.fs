namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel
open System.ComponentModel.DataAnnotations
open AdactaInternational.AdactaReportsShoppingBag.Model
open Newtonsoft.Json

type Product() =
    let propertyChanged = Event<PropertyChangedEventHandler, PropertyChangedEventArgs>()
    
    [<XmlAttribute("sigla")>]
    [<Required>]
    [<RegularExpression(@"^[A-Z0-9]{3}$")>]
    member val Code = "" with get, set

    [<XmlAttribute("nome")>]
    [<Required>]
    [<StringLength(100, MinimumLength = 1)>]
    member val Name = "" with get, set

    [<XmlAttribute("categoria")>]
    [<Required>]
    member val Category = "" with get, set

    [<XmlAttribute("produttore")>]
    [<Required>]
    member val Manufacturer = "" with get, set

    [<XmlAttribute("formato")>]
    [<Required>]
    member val Format = "" with get, set

    [<XmlAttribute("EAN")>]
    [<Required>]
    member val EAN = "" with get, set

    [<XmlAttribute("lotto")>]
    [<Required>]
    member val Batch = "" with get, set

    [<XmlAttribute("dataDiProduzione")>]
    [<Required>]
    member val ProductionDate = "" with get, set

    [<XmlAttribute("dataDiScadenza")>]
    [<Required>]
    member val ExpiryDate = "" with get, set

    [<XmlAttribute("stabilimentoDiProduzione")>]
    [<Required>]
    member val ProductionFacility = "" with get, set

    [<XmlElement("ingredienti")>]
    [<Required>]
    member val Ingredients = "" with get, set

    [<XmlArray("foto")>]
    [<XmlArrayItem("FotoProdotto")>]
    [<Required>]
    [<Length(1, 5)>]
    member val ProductPhotos: ProductPhoto array = [||] with get, set

    member val private _classification = ProductClassification.Unknown with get, set

    [<XmlIgnore>]
    member this.Classification 
        with get() = this._classification
        and set(value) = 
            if this._classification <> value then
                this._classification <- value
                propertyChanged.Trigger(this, PropertyChangedEventArgs("Classification"))
                propertyChanged.Trigger(this, PropertyChangedEventArgs("NavigationMenuItemIcon"))

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

    interface INotifyPropertyChanged with
        [<CLIEvent>]
        member _.PropertyChanged = propertyChanged.Publish
