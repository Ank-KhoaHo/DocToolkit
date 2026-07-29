namespace DocToolkit;

/// <summary>Thrown when a document conversion fails.</summary>
public sealed class DocumentConversionException : Exception
{
    public DocumentConversionException(string message) : base(message) { }
    public DocumentConversionException(string message, Exception inner) : base(message, inner) { }
}
