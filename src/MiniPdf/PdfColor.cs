namespace MiniPdf;

/// <summary>
/// Represents an RGB color for PDF text rendering.
/// Component values range from 0.0 (none) to 1.0 (full intensity).
/// </summary>
public readonly struct PdfColor : IEquatable<PdfColor>
{
    /// <summary>Red component (0.0–1.0).</summary>
    public float R { get; }

    /// <summary>Green component (0.0–1.0).</summary>
    public float G { get; }

    /// <summary>Blue component (0.0–1.0).</summary>
    public float B { get; }

    /// <summary>
    /// Creates a new PDF color from RGB components (0.0–1.0).
    /// </summary>
    public PdfColor(float r, float g, float b)
    {
        R = Math.Clamp(r, 0f, 1f);
        G = Math.Clamp(g, 0f, 1f);
        B = Math.Clamp(b, 0f, 1f);
    }

    /// <summary>
    /// Creates a PDF color from 0–255 byte RGB values.
    /// </summary>
    public static PdfColor FromRgb(byte r, byte g, byte b)
        => new(r / 255f, g / 255f, b / 255f);

    /// <summary>
    /// Creates a PDF color from a hex string (e.g. "FF0000" or "#FF0000").
    /// </summary>
    public static PdfColor FromHex(string hex)
    {
        if (string.IsNullOrEmpty(hex))
            return Black;

        hex = hex.TrimStart('#');

        if (hex.Length == 6 &&
            byte.TryParse(hex[0..2], System.Globalization.NumberStyles.HexNumber, null, out var r) &&
            byte.TryParse(hex[2..4], System.Globalization.NumberStyles.HexNumber, null, out var g) &&
            byte.TryParse(hex[4..6], System.Globalization.NumberStyles.HexNumber, null, out var b))
        {
            return FromRgb(r, g, b);
        }

        // ARGB format (8 chars) — skip alpha
        if (hex.Length == 8 &&
            byte.TryParse(hex[2..4], System.Globalization.NumberStyles.HexNumber, null, out r) &&
            byte.TryParse(hex[4..6], System.Globalization.NumberStyles.HexNumber, null, out g) &&
            byte.TryParse(hex[6..8], System.Globalization.NumberStyles.HexNumber, null, out b))
        {
            return FromRgb(r, g, b);
        }

        return Black;
    }

    /// <summary>Black (default text color).</summary>
    public static PdfColor Black => new(0, 0, 0);

    /// <summary>Red.</summary>
    public static PdfColor Red => new(1, 0, 0);

    /// <summary>Green.</summary>
    public static PdfColor Green => new(0, 0.5f, 0);

    /// <summary>Blue.</summary>
    public static PdfColor Blue => new(0, 0, 1);

    /// <summary>White.</summary>
    public static PdfColor White => new(1, 1, 1);

    /// <summary>Returns true if this color is black (0,0,0).</summary>
    public bool IsBlack => R == 0f && G == 0f && B == 0f;

    /// <inheritdoc />
    public bool Equals(PdfColor other) => R == other.R && G == other.G && B == other.B;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is PdfColor c && Equals(c);
    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(R, G, B);
    /// <summary>Equality operator.</summary>
    public static bool operator ==(PdfColor left, PdfColor right) => left.Equals(right);
    /// <summary>Inequality operator.</summary>
    public static bool operator !=(PdfColor left, PdfColor right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString() => $"PdfColor(R={R:F2}, G={G:F2}, B={B:F2})";
}
