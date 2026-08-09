using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Registers DocToolkit's DI-friendly services.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IHtmlToDocxConverter"/>, <see cref="IDocxToPdfConverter"/>,
    /// <see cref="IHtmlToPdfConverter"/>, <see cref="IDocxEditor"/>, <see cref="IWorkbookEditor"/>
    /// and <see cref="IPresentationEditor"/> as singletons - each wraps a stateless static class,
    /// so one shared instance is safe under concurrent use.
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

        return services;
    }
}
