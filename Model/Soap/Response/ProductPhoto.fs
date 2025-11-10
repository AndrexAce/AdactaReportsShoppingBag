namespace AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response

open System.Xml.Serialization
open System.ComponentModel.DataAnnotations
open CommunityToolkit.Mvvm.ComponentModel

type ProductPhoto() =
    inherit ObservableObject()

    let mutable _type = ""
    let mutable _photoUrl = ""

    [<XmlElement("tipo")>]
    [<Required>]
    member this.Type
        with get() = _type
        and set(value) = this.SetProperty(&_type, value) |> ignore

    [<XmlElement("urlFoto")>]
    [<Required>]
    member this.PhotoUrl
        with get() = _photoUrl
        and set(value) = this.SetProperty(&_photoUrl, value) |> ignore