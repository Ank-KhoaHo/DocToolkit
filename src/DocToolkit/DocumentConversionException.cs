namespace DocToolkit;

/// <summary>Thrown when a document conversion fails.</summary>
public sealed class DocumentConversionException : Exception
{
    /// <summary>Creates the exception with a message describing the failure.</summary>
    /// <param name="message">What could not be done.</param>
    public DocumentConversionException(string message) : base(message) { }

    /// <summary>Creates the exception wrapping the underlying failure.</summary>
    /// <param name="message">What could not be done.</param>
    /// <param name="inner">The library or I/O exception that caused it.</param>
    public DocumentConversionException(string message, Exception inner) : base(message, inner) { }
}
