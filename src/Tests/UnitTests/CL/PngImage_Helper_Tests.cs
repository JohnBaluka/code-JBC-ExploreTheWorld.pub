namespace JBC.ExploreTheWorld.UnitTests.CL;

public class PngImage_Helper_Tests
{
    // Minimal PNG header: 8-byte signature, IHDR length, "IHDR", then width/height (big-endian).
    private static byte[] MakePngHeader(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }.CopyTo(bytes, 0);
        bytes[11] = 13; // IHDR chunk length
        bytes[12] = (byte)'I'; bytes[13] = (byte)'H'; bytes[14] = (byte)'D'; bytes[15] = (byte)'R';
        bytes[16] = (byte)(width >> 24);  bytes[17] = (byte)(width >> 16);
        bytes[18] = (byte)(width >> 8);   bytes[19] = (byte)width;
        bytes[20] = (byte)(height >> 24); bytes[21] = (byte)(height >> 16);
        bytes[22] = (byte)(height >> 8);  bytes[23] = (byte)height;
        return bytes;
    }

    [Fact]
    public void TryGetPixelSize_ValidPngHeader_ReturnsWidthAndHeight()
    {
        var ok = PngImage_Helper.TryGetPixelSize(MakePngHeader(320, 213), out var width, out var height);

        ok.Should().BeTrue();
        width.Should().Be(320);
        height.Should().Be(213);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(new byte[0])]
    [InlineData(new byte[] { 1, 2, 3, 4 })]
    public void TryGetPixelSize_InvalidBytes_ReturnsFalse(byte[]? pngBytes)
    {
        PngImage_Helper.TryGetPixelSize(pngBytes, out _, out _).Should().BeFalse();
    }

    [Fact]
    public void GetWidthForHeight_ValidPng_PreservesAspectRatio()
    {
        // 320×160 → 2:1 aspect; display height 200 000 EMU → width 400 000 EMU.
        PngImage_Helper.GetWidthForHeight(MakePngHeader(320, 160), 200000).Should().Be(400000);
    }

    [Fact]
    public void GetWidthForHeight_InvalidPng_FallsBackTo3To2Ratio()
    {
        PngImage_Helper.GetWidthForHeight(null, 200000).Should().Be(300000);
    }
}
