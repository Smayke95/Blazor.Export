using PdfSharp.Fonts;
using S95.Blazor.Export.Fonts;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class ExportFontResolverTests
{
    private static readonly byte[] TrueTypeSignature = [0x00, 0x01, 0x00, 0x00];

    [Theory]
    [InlineData(false, "LiberationSans-Regular")]
    [InlineData(true, "LiberationSans-Bold")]
    public void ResolveTypeface_ReturnsEmbeddedFace(bool bold, string expectedFace)
    {
        var resolver = new ExportFontResolver(fallback: null);

        var info = resolver.ResolveTypeface(ExportFontResolver.FamilyName, bold, italic: false);

        Assert.NotNull(info);
        Assert.Equal(expectedFace, info.FaceName);
        Assert.False(info.MustSimulateBold);
    }

    [Theory]
    [InlineData("LiberationSans-Regular")]
    [InlineData("LiberationSans-Bold")]
    public void GetFont_ReturnsEmbeddedTrueTypeFont(string faceName)
    {
        var resolver = new ExportFontResolver(fallback: null);

        var font = resolver.GetFont(faceName);

        Assert.NotNull(font);
        Assert.Equal(TrueTypeSignature, font.Take(4));
    }

    [Fact]
    public void ResolveTypeface_WithUnknownFamily_UsesFallback()
    {
        var fallback = new FakeFontResolver();
        var resolver = new ExportFontResolver(fallback);

        var info = resolver.ResolveTypeface("Comic Sans MS", bold: false, italic: false);

        Assert.Equal(FakeFontResolver.FaceName, info?.FaceName);
        Assert.Equal(FakeFontResolver.Font, resolver.GetFont(FakeFontResolver.FaceName));
    }

    [Fact]
    public void ResolveTypeface_WithUnknownFamilyAndNoFallback_ReturnsNull()
    {
        var resolver = new ExportFontResolver(fallback: null);

        Assert.Null(resolver.ResolveTypeface("Comic Sans MS", bold: false, italic: false));
        Assert.Null(resolver.GetFont("ComicSans"));
    }

    [Fact]
    public void EnsureRegistered_RegistersResolverAsFallback()
    {
        ExportFontResolver.EnsureRegistered();
        ExportFontResolver.EnsureRegistered();

        Assert.IsType<ExportFontResolver>(GlobalFontSettings.FallbackFontResolver);
    }

    private class FakeFontResolver : IFontResolver
    {
        public const string FaceName = "Fake";

        public static readonly byte[] Font = [1, 2, 3];

        public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic) => new(FaceName);

        public byte[]? GetFont(string faceName) => faceName == FaceName ? Font : null;
    }
}