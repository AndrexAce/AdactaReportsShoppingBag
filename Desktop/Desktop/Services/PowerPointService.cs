using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using TextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class PowerPointService(INotificationService notificationService)
    : PowerPointComHandler, IPowerPointService
{
    public async Task CreateProductSlideshowAsync(Guid notificationId, ICollection<Product> products,
        string projectFolderPath, string projectName, string projectCode)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                CreateProductSlideshowInternal(notificationId, products, projectFolderPath, projectName, projectCode)));
    }

    private void CreateProductSlideshowInternal(Guid notificationId, ICollection<Product> products,
        string projectFolderPath, string projectName, string projectCode)
    {
        var productTotalCount = products.Count;
        var productCurrentCount = 0;

        // Track the COM classes to be released
        ExcelApplication? excelApp = null;
        Workbooks? workbooks = null;
        Workbook? productWorkbook = null;
        Sheets? productWorksheets = null;

        try
        {
            if (!Directory.Exists(Path.Combine(projectFolderPath, "Elaborazioni")))
                throw new DirectoryNotFoundException("Elaborazioni folder does not exist.");

            PowerPointApp = new PowerPointApplication
            {
                Visible = MsoTriState.msoTrue,
                DisplayAlerts = PpAlertLevel.ppAlertsNone
            };
            Presentations = PowerPointApp.Presentations;
            excelApp = new ExcelApplication
            {
                Visible = false,
                DisplayAlerts = false
            };
            workbooks = excelApp.Workbooks;

            // Extract the embedded template to a temporary file
            var templateStream = Assembly.GetExecutingAssembly()
                                     .GetManifestResourceStream(
                                         "AdactaInternational.AdactaReportsShoppingBag.Desktop.Assets.ProductTemplate.pptx")
                                 ?? throw new FileNotFoundException(
                                     "ProductTemplate.pptx embedded resource not found.");

            var tempTemplatePath = Path.Combine(Path.GetTempPath(), "ProductTemplate_temp.pptx");

            using (var fileStream = File.Create(tempTemplatePath))
            {
                templateStream.CopyTo(fileStream);
            }

            templateStream.Dispose();

            foreach (var product in products)
            {
                // Open the presentation file
                Presentation = Presentations.Open(tempTemplatePath);
                Slides = Presentation.Slides;

                // Open the product Excel file in read-only mode
                productWorkbook = workbooks.Open(
                    Path.Combine(projectFolderPath, "Elaborazioni", $"{product.DisplayName.Trim()}.xlsx"),
                    ReadOnly: true);
                productWorksheets = productWorkbook.Sheets;

                ProcessCoverSlideForProduct(product, Slides, projectName, projectCode);

                // Save the product presentation in the project folder
                var outputPath = Path.Combine(projectFolderPath, "Elaborazioni", $"{product.DisplayName.Trim()}.pptx");
                Presentation.SaveAs(outputPath);

                // Clean up for next iteration
                Marshal.ReleaseComObject(productWorksheets);
                productWorksheets = null;
                productWorkbook.Close(false);
                Marshal.ReleaseComObject(productWorkbook);
                productWorkbook = null;

                notificationService.UpdateProgressNotificationAsync(notificationId,
                    "Creazione file prodotti in corso...",
                    (uint)++productCurrentCount,
                    (uint)productTotalCount).GetAwaiter().GetResult();
            }

            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();

            // Clean up the temporary template file
            if (File.Exists(tempTemplatePath))
                File.Delete(tempTemplatePath);
        }
        catch (Exception e)
        {
            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();
            notificationService.ShowNotification("Elaborazione fallita",
                "Si è verificato un errore durante la creazione della presentazione del prodotto: " + e.Message);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (productWorksheets is not null) Marshal.ReleaseComObject(productWorksheets);
            if (productWorkbook is not null)
            {
                productWorkbook.Close(false);
                Marshal.ReleaseComObject(productWorkbook);
            }

            if (workbooks is not null) Marshal.ReleaseComObject(workbooks);
            if (excelApp is not null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }

    private static void ProcessCoverSlideForProduct(Product product, Slides productSlides, string projectName,
        string projectCode)
    {
        var coverSlide = productSlides[1];
        var shapes = coverSlide.Shapes;

        for (var i = shapes.Count; i >= 1; i--)
        {
            var shape = shapes[i];
            var shapeType = shape.Type;
            var shapeName = shape.Name;
            TextFrame? textFrame = null;
            TextRange? textRange = null;

            switch (shapeType, shapeName)
            {
                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Footer):
                    textFrame = shape.TextFrame;
                    textRange = textFrame.TextRange;
                    textRange.Text = $"{projectName} - job {projectCode} - {DateOnly.FromDateTime(DateTime.Now)}";

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Title):
                    textFrame = shape.TextFrame;
                    textRange = textFrame.TextRange;
                    textRange.Text = product.DisplayName;

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Photo):
                    var photoUrl = product.ProductPhotos.ElementAtOrDefault(0)?.PhotoUrl;

                    if (!string.IsNullOrEmpty(photoUrl))
                    {
                        var localImagePath = DownloadImageAsync(photoUrl).GetAwaiter().GetResult();

                        if (localImagePath is not null && File.Exists(localImagePath))
                        {
                            // Get placeholder dimensions and position
                            var left = shape.Left;
                            var top = shape.Top;

                            // Add the picture at the same position
                            var picture = shapes.AddPicture(
                                localImagePath,
                                MsoTriState.msoFalse,
                                MsoTriState.msoTrue,
                                left, top);

                            picture.ScaleHeight(0.2f, MsoTriState.msoTrue);
                            picture.ScaleWidth(0.2f, MsoTriState.msoTrue);

                            File.Delete(localImagePath);
                            Marshal.ReleaseComObject(picture);
                        }
                    }

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Logo):
                    if (!string.IsNullOrEmpty(product.Brand))
                    {
                        var localImagePath = LoadBrandImageAsync(product.Brand).GetAwaiter().GetResult();

                        if (localImagePath is not null && File.Exists(localImagePath))
                        {
                            // Get placeholder dimensions and position
                            var left = shape.Left;
                            var top = shape.Top;

                            // Add the picture at the same position
                            var picture = shapes.AddPicture(
                                localImagePath,
                                MsoTriState.msoFalse,
                                MsoTriState.msoTrue,
                                left, top);

                            picture.ScaleHeight(0.5f, MsoTriState.msoTrue);
                            picture.ScaleWidth(0.5f, MsoTriState.msoTrue);

                            File.Delete(localImagePath);
                            Marshal.ReleaseComObject(picture);
                        }
                    }

                    break;
            }

            if (textRange is not null) Marshal.ReleaseComObject(textRange);
            if (textFrame is not null) Marshal.ReleaseComObject(textFrame);
            Marshal.ReleaseComObject(shape);
        }

        Marshal.ReleaseComObject(shapes);
        Marshal.ReleaseComObject(coverSlide);
    }

    private static async Task<string?> DownloadImageAsync(string imageUrl)
    {
        try
        {
            using var httpClient = new HttpClient();
            var imageBytes = await httpClient.GetByteArrayAsync(imageUrl);

            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.jpg");
            await File.WriteAllBytesAsync(tempPath, imageBytes);

            return tempPath;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<string?> LoadBrandImageAsync(string brandName)
    {
        try
        {
            var productBrand = brandName.Trim().ToLower();

            var stream = Assembly.GetExecutingAssembly()
                             .GetManifestResourceStream(
                                 $"AdactaInternational.AdactaReportsShoppingBag.Desktop.Assets.{productBrand}.png")
                         ?? throw new FileNotFoundException(
                             $"{productBrand}.png embedded resource not found.");

            var tempPath = Path.Combine(Path.GetTempPath(), $"{productBrand}_temp.png");

            await using (var fileStream = File.Create(tempPath))
            {
                await stream.CopyToAsync(fileStream);
            }

            await stream.DisposeAsync();

            return tempPath;
        }
        catch
        {
            return null;
        }
    }

    private static class PlaceholderConstants
    {
        public const string Title = "Titolo";
        public const string Footer = "Footer";
        public const string Photo = "Foto";
        public const string Logo = "Logo";
        public const string SampleBase = "BaseCampione";
        public const string UserSampleBase = "BaseUser";
        public const string ProductSheetTable = "SchedaProdotto";
        public const string FrontPhoto = "Foto1";
        public const string RearPhoto = "Foto2";
        public const string UnpackagedPhoto = "Foto3";
        public const string ProductionEstabilishmentPhoto = "Foto4";
    }
}