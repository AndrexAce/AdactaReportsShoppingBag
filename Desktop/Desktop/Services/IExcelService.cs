using AdactaInternational.AdactaReportsShoppingBag.Model;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IExcelService
{
    public void CreateExcelClassesFile(ReportPrj project, string projectFolderPath);
}
