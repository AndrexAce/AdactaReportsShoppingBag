namespace AdactaInternational.AdactaReportsShoppingBag.Model

open AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response
open System.ComponentModel.DataAnnotations
open System.Collections.Generic

type ReportPrj =
    { [<Required>]
      [<RegularExpression(@"^[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}$")>]
      [<StringLength(11, MinimumLength = 7)>]
      Version: string

      [<Required>]
      [<StringLength(100, MinimumLength = 1)>]
      ProjectName: string

      [<Required>]
      [<RegularExpression(@"^[0-9]{2}\.[0-9]{3}[A-Z]?$")>]
      [<StringLength(7, MinimumLength = 6)>]
      ProjectCode: string

      [<Required>]
      [<MinLength(1)>]
      Products: IEnumerable<Product> }
