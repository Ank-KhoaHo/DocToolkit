using DocToolkit;

Console.WriteLine("Protection");
Console.WriteLine("==========");

// Every document below is produced by the library, so this sample needs no fixture.
byte[] statement = DocxToPdfConverter.Convert(
    DocxEditor.Create(new[] { DocxBlock.Paragraph("Statement for April") }));

byte[] workbook = WorkbookEditor.Create("Payroll", new[]
{
    new object?[] { "Name", "Net pay" },
    new object?[] { "A. Employee", 2810.44 },
});

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n1. A PDF nobody can open without the password");

#region protect-pdf-user-password
// A USER password is required to OPEN the file. This is the one enforced by cryptography.
byte[] locked = PdfEditor.Protect(statement, new PdfProtection { UserPassword = "s3cret" });

// Every other PdfEditor operation refuses an encrypted document, so take the protection off first.
// If the document also has an owner password, THAT is the one Unprotect needs.
byte[] opened = PdfEditor.Unprotect(locked, "s3cret");
#endregion

Console.WriteLine($"   original     : {statement.Length,7:N0} bytes, {PdfEditor.PageCount(statement)} page(s)");
Console.WriteLine($"   protected    : {locked.Length,7:N0} bytes");

try
{
    PdfEditor.PageCount(locked);
    Console.WriteLine("   reading it   : SUCCEEDED - which would mean it is not really protected");
}
catch (DocumentConversionException)
{
    Console.WriteLine("   reading it   : refused, as it should be");
}

Console.WriteLine($"   unprotected  : {PdfEditor.PageCount(opened)} page(s) again");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n2. The trap: an owner password is NOT a lock");

#region protect-owner-password-is-not-a-lock
// An OWNER password sets the permissions and nothing more. The file still opens for anyone - the
// flags are a request a reader is asked to honour, not something it is prevented from ignoring.
byte[] restricted = PdfEditor.Protect(statement, new PdfProtection
{
    OwnerPassword = "admin",
    AllowPrinting = false,
    AllowCopying = false,
});

// ...and here is the proof, in the same breath: no password was supplied, and the text comes out.
IReadOnlyList<string> readableAnyway = PdfEditor.ExtractText(restricted);
#endregion

Console.WriteLine($"   encrypted?   : yes ({restricted.Length:N0} bytes)");
Console.WriteLine($"   text read with NO password: \"{readableAnyway[0].Trim()}\"");
Console.WriteLine("   ^ that is the PDF specification working as designed, not a defect.");
Console.WriteLine("     If the content must not be READ, set UserPassword.");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n3. The same three members on every Office format");

#region protect-office-formats
// DocxEditor, WorkbookEditor and PresentationEditor each carry Protect, Unprotect and IsProtected.
byte[] lockedWorkbook = WorkbookEditor.Protect(workbook, "s3cret");

// IsProtected reads the file signature, so it needs no password: it answers "will the other
// members refuse this?" for a file you were handed and know nothing about.
bool needsPassword = WorkbookEditor.IsProtected(lockedWorkbook);

byte[] openedWorkbook = WorkbookEditor.Unprotect(lockedWorkbook, "s3cret");
#endregion

Console.WriteLine($"   IsProtected(locked)  : {needsPassword}");
Console.WriteLine($"   IsProtected(original): {WorkbookEditor.IsProtected(workbook)}");
Console.WriteLine($"   cell after unprotect : {WorkbookEditor.ReadCell(openedWorkbook, "Payroll", "B2")}");

// An encrypted Office file is not a package any more - the ZIP is sealed inside a compound file.
// That is why the ordinary members refuse it rather than returning something half-read.
Console.WriteLine($"   locked starts with   : {(char)lockedWorkbook[0]}{(char)lockedWorkbook[1]} "
                  + "(a compound file, not the PK of a ZIP)");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n4. Two failures worth telling apart");

#region protect-distinguish-failures
// A wrong password and a file that was never encrypted are different problems, and only one of
// them is fixed by asking the user to type it again.
foreach (var (label, bytes, password) in new[]
{
    ("wrong password    ", lockedWorkbook, "not-the-password"),
    ("never encrypted   ", workbook, "s3cret"),
})
{
    try
    {
        WorkbookEditor.Unprotect(bytes, password);
        Console.WriteLine($"   {label}: opened - unexpected");
    }
    catch (DocumentConversionException ex)
    {
        Console.WriteLine($"   {label}: {ex.Message}");
    }
}
#endregion

string outputPath = Path.Join(AppContext.BaseDirectory, "protected.pdf");
await File.WriteAllBytesAsync(outputPath, locked);
Console.WriteLine($"\nWrote {outputPath} (opens with the password s3cret)");

Console.WriteLine("\nDone.");
