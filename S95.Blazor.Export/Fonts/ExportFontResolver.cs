using PdfSharp.Fonts;

namespace S95.Blazor.Export.Fonts;

internal class ExportFontResolver : IFontResolver
{
    public const string FamilyName = "Liberation Sans";

    private const string RegularFace = "LiberationSans-Regular";
    private const string BoldFace = "LiberationSans-Bold";

    private static readonly object RegistrationLock = new();

    private readonly IFontResolver? Fallback;

    public ExportFontResolver(IFontResolver? fallback)
    {
        Fallback = fallback;
    }

    public static void EnsureRegistered()
    {
        lock (RegistrationLock)
        {
            if (GlobalFontSettings.FallbackFontResolver is ExportFontResolver)
                return;

            GlobalFontSettings.FallbackFontResolver = new ExportFontResolver(GlobalFontSettings.FallbackFontResolver);
        }
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        if (!string.Equals(familyName, FamilyName, StringComparison.OrdinalIgnoreCase))
            return Fallback?.ResolveTypeface(familyName, bold, italic);

        return new FontResolverInfo(bold ? BoldFace : RegularFace, mustSimulateBold: false, mustSimulateItalic: italic);
    }

    public byte[]? GetFont(string faceName)
    {
        if (faceName is not (RegularFace or BoldFace))
            return Fallback?.GetFont(faceName);

        using var stream = typeof(ExportFontResolver).Assembly.GetManifestResourceStream($"S95.Blazor.Export.Fonts.{faceName}.ttf")
            ?? throw new InvalidOperationException($"Embedded font '{faceName}' was not found.");

        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }
}