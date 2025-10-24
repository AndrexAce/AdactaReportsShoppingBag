namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Project

open System.ComponentModel.DataAnnotations

type ReportPrj = {
    [<Required>]
    [<RegularExpression(@"^\d{1,}\.\d{1,}\.\d{1,}$")>]
    [<StringLength(8, MinimumLength = 5)>]
    Version: string

    [<Required>]
    [<StringLength(100, MinimumLength = 1)>]
    ProjectName: string

    [<Required>]
    [<RegularExpression(@"^\d{2}\.\d{3}[a-zA-Z]{0,1}$")>]
    [<StringLength(7, MinimumLength = 6)>]
    ProjectCode: string
}