namespace AdactaInternational.AdactaReportsShoppingBag.Model.Project

open System.ComponentModel
open System.ComponentModel.DataAnnotations
open System.Runtime.CompilerServices

type ReportPrj() =
    // Define the PropertyChanged event
    let propertyChanged = Event<PropertyChangedEventHandler, PropertyChangedEventArgs>()
    
    // Backing fields for properties
    let mutable version = ""
    let mutable projectName = ""
    let mutable projectCode = ""
    
    // Helper method to raise property changed
    member private this.OnPropertyChanged([<CallerMemberName>] ?propertyName: string) =
        let name = defaultArg propertyName ""
        propertyChanged.Trigger(this, PropertyChangedEventArgs(name))
    
    // Version property with validation attributes
    [<Required>]
    [<RegularExpression(@"^\d{1,}\.\d{1,}\.\d{1,}$")>]
    [<StringLength(8, MinimumLength = 5)>]
    member this.Version
        with get() = version
        and set(value) =
            if version <> value then
                version <- value
                this.OnPropertyChanged()
    
    // ProjectName property with validation attributes
    [<Required>]
    [<StringLength(100, MinimumLength = 1)>]
    member this.ProjectName
        with get() = projectName
        and set(value) =
            if projectName <> value then
                projectName <- value
                this.OnPropertyChanged()
    
    // ProjectCode property with validation attributes
    [<Required>]
    [<RegularExpression(@"^\d{2}\.\d{3}[a-zA-Z]{0,1}$")>]
    [<StringLength(7, MinimumLength = 6)>]
    member this.ProjectCode
        with get() = projectCode
        and set(value) =
            if projectCode <> value then
                projectCode <- value
                this.OnPropertyChanged()
    
    // Expose the PropertyChanged event as a CLI event
    [<CLIEvent>]
    member this.PropertyChanged = propertyChanged.Publish
    
    // Implement INotifyPropertyChanged interface
    interface INotifyPropertyChanged with
        member this.add_PropertyChanged(handler) = this.PropertyChanged.AddHandler(handler)
        member this.remove_PropertyChanged(handler) = this.PropertyChanged.RemoveHandler(handler)