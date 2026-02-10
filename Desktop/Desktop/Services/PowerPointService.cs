using System.Reflection;
using System.Runtime.InteropServices;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
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
        Presentation? presentation = null;
        Slides? slides = null;

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

            // Extract the embedded template to a temporary file
            var tempTemplatePath = Path.Combine(Path.GetTempPath(), "ProductTemplate_temp.pptx");

            using (var templateStream = Assembly.GetExecutingAssembly()
                                            .GetManifestResourceStream(
                                                "AdactaInternational.AdactaReportsShoppingBag.Desktop.Assets.ProductTemplate.pptx")
                                        ?? throw new FileNotFoundException(
                                            "ProductTemplate.pptx embedded resource not found."))
            {
                using (var fileStream = File.Create(tempTemplatePath))
                {
                    templateStream.CopyTo(fileStream);
                }
            }

            foreach (var product in products)
            {
                // Open the presentation file
                presentation = Presentations.Open(tempTemplatePath);
                slides = presentation.Slides;

                ProcessCoverSlideForProduct(product, slides, projectName, projectCode);
                ProcessFirstSlideForProduct(product, slides, projectName, projectCode);

                // Save the product presentation in the project folder
                var outputPath = Path.Combine(projectFolderPath, "Elaborazioni", $"{product.DisplayName.Trim()}.pptx");
                presentation.SaveAs(outputPath);

                notificationService.UpdateProgressNotificationAsync(notificationId,
                    "Creazione file prodotti in corso...",
                    (uint)++productCurrentCount,
                    (uint)productTotalCount).GetAwaiter().GetResult();

                if (slides is not null) Marshal.ReleaseComObject(slides);
                slides = null;
                presentation.Close();
                Marshal.ReleaseComObject(presentation);
                presentation = null;
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
            if (slides is not null) Marshal.ReleaseComObject(slides);

            if (presentation is not null)
            {
                presentation.Close();
                Marshal.ReleaseComObject(presentation);
            }
        }
    }

    private static void ProcessCoverSlideForProduct(Product product, Slides productSlides, string projectName,
        string projectCode)
    {
        var coverSlide = productSlides[1];
        var shapes = coverSlide.Shapes;

        foreach (Shape? shape in shapes)
        {
            var shapeType = shape?.Type;
            var shapeName = shape?.Name;
            TextFrame? textFrame = null;
            TextRange? textRange = null;

            switch (shapeType, shapeName)
            {
                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Footer):
                    textFrame = shape?.TextFrame;
                    textRange = textFrame?.TextRange;
                    textRange?.Text = $"{projectName} - job {projectCode} - {DateOnly.FromDateTime(DateTime.Now)}";

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Title):
                    textFrame = shape?.TextFrame;
                    textRange = textFrame?.TextRange;
                    textRange?.Text = product.ProductName;

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Photo):
                    var photoUrl = product.ProductPhotos.ElementAtOrDefault(0)?.PhotoUrl;

                    if (!string.IsNullOrEmpty(photoUrl))
                    {
                        var localImagePath = DownloadImageAsync(photoUrl).GetAwaiter().GetResult();

                        if (localImagePath is not null && File.Exists(localImagePath))
                        {
                            // Get placeholder dimensions and position
                            var left = shape?.Left;
                            var top = shape?.Top;

                            // Add the picture at the same position
                            var picture = shapes.AddPicture(
                                localImagePath,
                                MsoTriState.msoFalse,
                                MsoTriState.msoTrue,
                                left ?? 0, top ?? 0);

                            picture.ScaleHeight(0.2f, MsoTriState.msoTrue);
                            picture.ScaleWidth(0.2f, MsoTriState.msoTrue);
                            picture.Name = PlaceholderConstants.Photo;

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
                            var left = shape?.Left;
                            var top = shape?.Top;

                            // Add the picture at the same position
                            var picture = shapes.AddPicture(
                                localImagePath,
                                MsoTriState.msoFalse,
                                MsoTriState.msoTrue,
                                left ?? 0, top ?? 0);

                            picture.ScaleHeight(0.5f, MsoTriState.msoTrue);
                            picture.ScaleWidth(0.5f, MsoTriState.msoTrue);
                            picture.Name = PlaceholderConstants.Logo;

                            File.Delete(localImagePath);
                            Marshal.ReleaseComObject(picture);
                        }
                    }

                    break;
            }

            if (textRange is not null) Marshal.ReleaseComObject(textRange);
            if (textFrame is not null) Marshal.ReleaseComObject(textFrame);
            if (shape is not null) Marshal.ReleaseComObject(shape);
        }

        Marshal.ReleaseComObject(shapes);
        Marshal.ReleaseComObject(coverSlide);
    }

    private static void ProcessFirstSlideForProduct(Product product, Slides productSlides, string projectName,
        string projectCode)
    {
        var coverSlide = productSlides[2];
        var shapes = coverSlide.Shapes;

        foreach (Shape? shape in shapes)
        {
            var shapeType = shape?.Type;
            var shapeName = shape?.Name;
            TextFrame? textFrame = null;
            TextRange? textRange = null;

            switch (shapeType, shapeName)
            {
                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Footer):
                    textFrame = shape?.TextFrame;
                    textRange = textFrame?.TextRange;
                    textRange?.Text = $"{projectName} - job {projectCode} - {DateOnly.FromDateTime(DateTime.Now)}";

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Title):
                    textFrame = shape?.TextFrame;
                    textRange = textFrame?.TextRange;
                    textRange?.Text = product.ProductName;

                    break;

                case (MsoShapeType.msoPlaceholder, PlaceholderConstants.Logo):
                    if (!string.IsNullOrEmpty(product.Brand))
                    {
                        var localImagePath = LoadBrandImageAsync(product.Brand).GetAwaiter().GetResult();

                        if (localImagePath is not null && File.Exists(localImagePath))
                        {
                            // Get placeholder dimensions and position
                            var left = shape?.Left;
                            var top = shape?.Top;

                            // Add the picture at the same position
                            var picture = shapes.AddPicture(
                                localImagePath,
                                MsoTriState.msoFalse,
                                MsoTriState.msoTrue,
                                left ?? 0, top ?? 0);

                            picture.Name = PlaceholderConstants.Logo;

                            File.Delete(localImagePath);
                            Marshal.ReleaseComObject(picture);
                        }
                    }

                    break;
            }

            if (textRange is not null) Marshal.ReleaseComObject(textRange);
            if (textFrame is not null) Marshal.ReleaseComObject(textFrame);
            if (shape is not null) Marshal.ReleaseComObject(shape);
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