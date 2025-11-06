namespace AdactaInternational.AdactaReportsShoppingBag.Model

open AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response
open System.ComponentModel
open System.ComponentModel.DataAnnotations
open System.Runtime.CompilerServices
open System.Collections.Generic

type ReportPrj() =
    let propertyChanged = Event<PropertyChangedEventHandler, PropertyChangedEventArgs>()

    let mutable version = ""
    let mutable projectName = ""
    let mutable projectCode = ""
    let mutable products: IEnumerable<Product> = []

    member private this.OnPropertyChanged([<CallerMemberName>] ?propertyName: string) =
        let name = defaultArg propertyName ""
        propertyChanged.Trigger(this, PropertyChangedEventArgs(name))

    [<Required>]
    [<RegularExpression(@"^[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}$")>]
    [<StringLength(11, MinimumLength = 7)>]
    member this.Version
        with get () = version
        and set (value) =
            if version <> value then
                version <- value
                this.OnPropertyChanged()

    [<Required>]
    [<StringLength(100, MinimumLength = 1)>]
    member this.ProjectName
        with get () = projectName
        and set (value) =
            if projectName <> value then
                projectName <- value
                this.OnPropertyChanged()

    [<Required>]
    [<RegularExpression(@"^[0-9]{2}\.[0-9]{3}[A-Z]?$")>]
    [<StringLength(7, MinimumLength = 6)>]
    member this.ProjectCode
        with get () = projectCode
        and set (value) =
            if projectCode <> value then
                projectCode <- value
                this.OnPropertyChanged()

    [<Required>]
    [<MinLength(1)>]
    member this.Products
        with get () = products
        and set (value) =
            if products <> value then
                products <- value
                this.OnPropertyChanged()

    [<CLIEvent>]
    member this.PropertyChanged = propertyChanged.Publish

    interface INotifyPropertyChanged with
        member this.add_PropertyChanged(handler) =
            this.PropertyChanged.AddHandler(handler)

        member this.remove_PropertyChanged(handler) =
            this.PropertyChanged.RemoveHandler(handler)
