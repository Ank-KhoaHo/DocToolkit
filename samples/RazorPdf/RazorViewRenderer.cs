using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Abstractions;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.Razor;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.AspNetCore.Routing;

namespace RazorPdf;

/// <summary>
/// Renders a Razor view to a string, outside the request pipeline.
/// </summary>
/// <remarks>
/// This is the piece people expect to find in the box and never do: MVC renders views straight to
/// the response, so getting the HTML as a <see cref="string"/> means driving the view engine
/// yourself. It is about twenty lines, and it is the same twenty lines whatever you do with the
/// result — email bodies, snapshots, or, here, a PDF.
///
/// The <see cref="ActionContext"/> is a stand-in. Rendering needs one, but nothing in a document
/// template reads the request, so an empty <c>DefaultHttpContext</c> carrying the application's
/// <c>RequestServices</c> is enough. That is also what lets this run from a background job with no
/// request in flight at all.
/// </remarks>
public sealed class RazorViewRenderer(
    IRazorViewEngine viewEngine,
    ITempDataProvider tempDataProvider,
    IServiceProvider services)
{
    public async Task<string> RenderAsync<TModel>(string viewPath, TModel model)
    {
        var actionContext = new ActionContext(
            new DefaultHttpContext { RequestServices = services },
            new RouteData(),
            new ActionDescriptor());

        // GetView by explicit path rather than FindView by name: FindView resolves through the
        // route values of the action being executed, and there is no action here.
        var result = viewEngine.GetView(executingFilePath: null, viewPath, isMainPage: true);
        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"View '{viewPath}' was not found. Searched: {string.Join(", ", result.SearchedLocations)}");
        }

        await using var output = new StringWriter();

        var viewContext = new ViewContext(
            actionContext,
            result.View,
            new ViewDataDictionary<TModel>(new EmptyModelMetadataProvider(), new ModelStateDictionary())
            {
                Model = model,
            },
            new TempDataDictionary(actionContext.HttpContext, tempDataProvider),
            output,
            new HtmlHelperOptions());

        await result.View.RenderAsync(viewContext);

        return output.ToString();
    }
}
