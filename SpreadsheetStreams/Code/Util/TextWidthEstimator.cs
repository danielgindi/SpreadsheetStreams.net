using System;
using System.Drawing;
using System.Globalization;

#nullable enable

namespace SpreadsheetStreams.Util
{
    public static class TextWidthEstimator
    {
        /// <summary>
        /// Estimates Excel column width for a single line of text
        /// (Calibri 11 default). Newlines should be handled externally.
        /// </summary>
        public static float MeasureText(
            string text,
            bool multiLine,
            float emojiWidth = 2.0f,
            float padding = 0.9f,
            float minWidth = 0.0f,
            float maxWidth = 255.0f)
        {
            if (string.IsNullOrEmpty(text))
                return Math.Min(maxWidth, Math.Max(minWidth, padding));

            float multiSize = 0.0f;
            float size = 0.0f;

            bool pendingHigh = false;
            foreach (var c in text)
            {
                if (!pendingHigh)
                {
                    if (!char.IsHighSurrogate(c))
                    {
                        // Lone low surrogate: treat like emoji/symbol
                        if (char.IsLowSurrogate(c))
                        {
                            size += emojiWidth;
                            continue;
                        }

                        size += GlyphWidth(c);

                        if (size >= maxWidth)
                            break;

                        continue;
                    }

                    if (c == '\n')
                    {
                        multiSize = Math.Max(multiSize, size);
                        size = 0;
                    }

                    // High surrogate: count as emoji and swallow the low surrogate if it comes next
                    size += emojiWidth;
                    pendingHigh = true;
                    continue;
                }

                // We already counted emojiWidth for the previous high surrogate.
                // If current char is its matching low surrogate, skip it.
                if (char.IsLowSurrogate(c))
                {
                    pendingHigh = false;
                    continue;
                }

                // High surrogate not followed by low surrogate:
                // process this char normally (and it might itself be another high surrogate)
                pendingHigh = false;

                if (char.IsHighSurrogate(c))
                {
                    size += emojiWidth;
                    pendingHigh = true;
                }
                else if (char.IsLowSurrogate(c))
                {
                    size += emojiWidth;
                }
                else
                {
                    size += GlyphWidth(c);
                }

                if (size >= maxWidth)
                    break;
            }

            size = Math.Max(size, multiSize);

            return Math.Min(maxWidth, Math.Max(minWidth, size + padding));
        }

        public static float GlyphWidth(char ch)
        {
            // Whitespace
            if (ch == ' ' || ch == '\u00A0') return 0.55f;

            // Hebrew diacritics (niqqud / cantillation)
            if (IsHebrewDiacritic(ch)) return 0.05f;

            // Hebrew punctuation
            if (ch == '\u05F3') return 0.45f; // ׳
            if (ch == '\u05F4') return 0.65f; // ״

            // Hebrew letters (Calibri 11-ish buckets)
            if (IsHebrewLetter(ch))
                return HebrewLetterWidth(ch);

            // Latin letters
            if (IsLatinLetter(ch))
            {
                if (ch is 'i' or 'l' or 'I' or 'j' or 't' or 'f' or 'r') return 0.55f;
                if (ch is 'M' or 'W' or '@' or '#' or '%' or '&') return 1.45f;
                if (char.IsUpper(ch)) return 1.10f;
                return 0.95f;
            }

            // Digits
            if (ch >= '0' && ch <= '9') return 0.95f;

            // Punctuation / symbols
            if (ch is '|' or '!' or ':' or '\'' or '"' or '.' or ',' or ';') return 0.45f;
            if (ch is '-' or '–' or '—' or '_') return 0.85f;
            if (ch is '(' or ')' or '[' or ']' or '{' or '}') return 0.75f;
            if (ch is '=' or '+' or '*' or '/' or '\\') return 0.95f;

            // Currency
            if (ch is '₪' or '$' or '€' or '£' or '¥') return 1.15f;

            // Unicode fallback (BMP)
            return CharUnicodeInfo.GetUnicodeCategory(ch) switch
            {
                UnicodeCategory.NonSpacingMark => 0.05f,
                UnicodeCategory.SpacingCombiningMark => 0.20f,
                UnicodeCategory.DecimalDigitNumber => 0.95f,
                UnicodeCategory.UppercaseLetter => 1.10f,
                UnicodeCategory.LowercaseLetter => 1.00f,
                UnicodeCategory.OtherLetter => 1.05f,
                UnicodeCategory.CurrencySymbol => 1.15f,
                UnicodeCategory.MathSymbol => 1.10f,
                UnicodeCategory.OtherSymbol => 1.20f,
                _ => 1.00f
            };
        }

        private static float HebrewLetterWidth(char c)
        {
            return c switch
            {
                // Very narrow
                'י' or 'ו' or 'ן' => 0.75f,

                // Narrow finals
                'ך' => 0.80f,

                // Slightly narrow
                'ר' or 'ז' or 'ג' => 0.90f,

                // Wide / blocky
                'מ' or 'ם' or 'ש' or 'ת' or 'ב' => 1.10f,

                // Slightly wide
                'ה' or 'ח' or 'כ' => 1.05f,

                _ => 1.00f
            };
        }

        private static bool IsLatinLetter(char ch) =>
            (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

        private static bool IsHebrewLetter(char ch) =>
            ch >= '\u05D0' && ch <= '\u05EA'; // א..ת

        private static bool IsHebrewDiacritic(char ch)
        {
            int cp = ch;
            return (cp >= 0x0591 && cp <= 0x05BD) ||
                   cp == 0x05BF ||
                   (cp >= 0x05C1 && cp <= 0x05C2) ||
                   (cp >= 0x05C4 && cp <= 0x05C5) ||
                   cp == 0x05C7;
        }
    }
}
