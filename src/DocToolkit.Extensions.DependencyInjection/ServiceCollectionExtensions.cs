using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Registers DocToolkit's DI-friendly services.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers every injectable interface as a singleton - each wraps a stateless static
    /// class, so one shared instance is safe under concurrent use.
    ///
    /// <see cref="DocToolkitOptions"/> is consumed through <c>IOptionsMonitor</c> and read on
    /// every call, so a configuration reload takes effect without restarting the process.
    /// </summary>
    /// <param name="services">The service collection to add to.</param>
    /// <param name="configure">
    /// Configures <see cref="DocToolkitOptions"/>. Leave null to keep every default -
    /// <see cref="DocToolkitOptions.AllowRemoteImageDownload"/> stays <c>false</c>, so nothing is
    /// fetched and <see cref="DocToolkitOptions.RemoteImage"/> never comes into play.
    /// </param>
    public static IServiceCollection AddDocToolkit(
        this IServiceCollection services, Action<DocToolkitOptions>? configure = null)
    {
        services.AddOptions<DocToolkitOptions>();
        if (configure is not null) services.Configure(configure);

        services.TryAddSingleton<IHtmlToDocxConverter, HtmlToDocxConverterService>();
        services.TryAddSingleton<IDocxToPdfConverter, DocxToPdfConverterService>();
        services.TryAddSingleton<IHtmlToPdfConverter, HtmlToPdfConverterService>();
        services.TryAddSingleton<IDocxEditor, DocxEditorService>();
        services.TryAddSingleton<IWorkbookEditor, WorkbookEditorService>();
        services.TryAddSingleton<IPresentationEditor, PresentationEditorService>();
        services.TryAddSingleton<IXlsxToPdfConverter, XlsxToPdfConverterService>();
        services.TryAddSingleton<IPptxToPdfConverter, PptxToPdfConverterService>();
        services.TryAddSingleton<IDocxToHtmlConverter, DocxToHtmlConverterService>();
        services.TryAddSingleton<IDocxToMarkdownConverter, DocxToMarkdownConverterService>();
        services.TryAddSingleton<IPdfEditor, PdfEditorService>();
        services.TryAddSingleton<IMarkdownToDocxConverter, MarkdownToDocxConverterService>();
        services.TryAddSingleton<IMarkdownToPdfConverter, MarkdownToPdfConverterService>();
        services.TryAddSingleton<IXlsxToCsvConverter, XlsxToCsvConverterService>();
        services.TryAddSingleton<IXlsxToHtmlConverter, XlsxToHtmlConverterService>();

        return services;
    }
}
