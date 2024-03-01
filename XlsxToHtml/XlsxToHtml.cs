using ClosedXML.Excel;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxToHtml
{
    /// <summary>
    /// Converts an Excel (.xlsx) file to HTML format.
    /// </summary>
    public class XlsxToHtml
    {
        /// <summary>
        /// Default format string for HTML.
        /// </summary>
        public const string DefaultFormatString = @"<html xmlns:o=""urn:schemas-microsoft-com:office: office""
xmlns:x=""urn:schemas-microsoft-com:office:excel""
xmlns=""http://www.w3.org/TR/REC-html40"">";

        /// <summary>
        /// Default string for HTML head section.
        /// </summary>
        public const string DefaultHeadString = @"<head>
<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>
<style>
body {
    margin: 0;
    padding: 0;
}
table {
    width: 100%;
    table-layout: fixed;
    border-collapse: collapse;
}
td {
    padding-top:1px;
	padding-right:1px;
	padding-left:1px;
    text-align: left;
    vertical-align: bottom;
    padding: 0px;
    cellColor: black;
    background-cellColor: transparent;
    border-width: 1px;
    border-style: solid;
    border-cellColor: lightgray;
    border-collapse: collapse;
    white-space: nowrap;
}
</style>
</head>";

        /// <summary>
        /// Converts the specified Excel file to HTML format.
        /// </summary>
        /// <param name="xlsxFilePath">The path to the Excel (.xlsx) file.</param>
        /// <returns>The HTML representation of the Excel file.</returns>
        public static string Convert(string xlsxFilePath)
        {
            StringBuilder htmlBuilder = new StringBuilder();
            using (var workbook = new XLWorkbook(xlsxFilePath))
            {
                IXLTheme theme = new XLThemeOffice();
                var worksheet = workbook.Worksheet(1); // 첫 번째 워크시트를 선택

                htmlBuilder.AppendLine(DefaultFormatString);
                htmlBuilder.AppendLine(DefaultHeadString);
                htmlBuilder.AppendLine("<body>");
                htmlBuilder.AppendLine("<div align=center>");
                htmlBuilder.AppendLine("<table>");

                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.IsHidden) continue;
                    if (row.RowNumber() >= 1048576) break;

                    htmlBuilder.AppendLine("<tr>");

                    foreach (var col in worksheet.ColumnsUsed(XLCellsUsedOptions.All))
                    {
                        if (col.IsHidden) continue;
                        if (col.ColumnNumber() >= 16384) break;

                        var cell = worksheet.Cell(row.RowNumber(), col.ColumnNumber());
                        if (cell.IsMerged())
                        {
                            if (cell.Address.RowNumber != cell.MergedRange().FirstCell().Address.RowNumber) continue;
                            if (cell.Address.ColumnNumber != cell.MergedRange().FirstCell().Address.ColumnNumber) continue;
                            htmlBuilder.Append(EncodeCellStyle(cell, cell.MergedRange(), theme));
                        }
                        else
                        {
                            htmlBuilder.Append(EncodeCellStyle(cell, theme));
                        }

                        try
                        {
                            string value = cell.Value.ToString();
                            htmlBuilder.Append(EncodeCellValue(value));
                        }
                        catch
                        {
                            Console.WriteLine($"Exception:cell value not supported:{cell.Address}");
                        }

                        htmlBuilder.AppendLine("</td>");
                    }

                    htmlBuilder.AppendLine("</tr>");
                }

                htmlBuilder.AppendLine("</table>");
                htmlBuilder.AppendLine("</div>");
                htmlBuilder.AppendLine("</body>");
                htmlBuilder.AppendLine("</html>");
            }

            return htmlBuilder.ToString();
        }

        /// <summary>
        /// Processes the cell style for a non-merged cell.
        /// </summary>
        /// <param name="currentCell">The current cell being processed.</param>
        /// <param name="theme">The theme to use for resolving cell styles.</param>
        /// <returns>The HTML representation of the cell style.</returns>
        private static string EncodeCellStyle(IXLCell currentCell, IXLTheme theme)
        {
            StringBuilder styleBuilder = new StringBuilder();
            styleBuilder.Append(PropertyToStyle("border-top", currentCell.Style.Border.TopBorder, currentCell.Style.Border.TopBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-bottom", currentCell.Style.Border.BottomBorder, currentCell.Style.Border.BottomBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-left", currentCell.Style.Border.LeftBorder, currentCell.Style.Border.LeftBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-right", currentCell.Style.Border.RightBorder, currentCell.Style.Border.RightBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("background-color", currentCell.Style.Fill.BackgroundColor, theme));
            styleBuilder.Append(PropertyToStyle("color", currentCell.Style.Font.FontColor, theme));
            styleBuilder.Append(PropertyToStyle("text-align", currentCell.Style.Alignment.Horizontal));
            styleBuilder.Append(PropertyToStyle("vertical-align", currentCell.Style.Alignment.Vertical));
            styleBuilder.Append(PropertyToStyle("font-size", (int)currentCell.Style.Font.FontSize, 11));
            styleBuilder.Append(PropertyToStyle("font-family", currentCell.Style.Font.FontName, "serif"));
            styleBuilder.Append(PropertyToStyle("font-style", currentCell.Style.Font.Italic ? "italic" : "normal", "normal"));
            styleBuilder.Append(PropertyToStyle("font-weight", currentCell.Style.Font.Bold ? "bold" : "normal", "normal"));
            styleBuilder.Append(PropertyToStyle("white-space", currentCell.Style.Alignment.WrapText ? "no-wrap" : "", string.Empty));
            styleBuilder.Append(PropertyToStyle("height", GetRowHeight(currentCell), 0));
            styleBuilder.Append(PropertyToStyle("width", GetColumnWidth(currentCell), 0));
            styleBuilder.Append(PropertyToStyle("transform", currentCell.Style.Alignment.TextRotation, 0));
            return $"<td style=\"{styleBuilder.ToString().TrimEnd()}\" eth-cell=\"{currentCell.Address}\">";
        }

        /// <summary>
        /// Processes the cell style for a merged cell.
        /// </summary>
        /// <param name="currentCell">The current cell being processed.</param>
        /// <param name="range">The range of merged cells.</param>
        /// <param name="theme">The theme to use for resolving cell styles.</param>
        /// <returns>The HTML representation of the cell style.</returns>
        private static string EncodeCellStyle(IXLCell currentCell, IXLRange range, IXLTheme theme)
        {
            StringBuilder styleBuilder = new StringBuilder();
            var lastCell = range.LastCell();
            styleBuilder.Append(PropertyToStyle("border-top", currentCell.Style.Border.TopBorder, currentCell.Style.Border.TopBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-bottom", currentCell.Style.Border.BottomBorder, currentCell.Style.Border.BottomBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-left", currentCell.Style.Border.LeftBorder, currentCell.Style.Border.LeftBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("border-right", lastCell.Style.Border.RightBorder, lastCell.Style.Border.RightBorderColor, theme));
            styleBuilder.Append(PropertyToStyle("background-color", currentCell.Style.Fill.BackgroundColor, theme));
            styleBuilder.Append(PropertyToStyle("color", currentCell.Style.Font.FontColor, theme));
            styleBuilder.Append(PropertyToStyle("text-align", currentCell.Style.Alignment.Horizontal));
            styleBuilder.Append(PropertyToStyle("vertical-align", currentCell.Style.Alignment.Vertical));
            styleBuilder.Append(PropertyToStyle("font-size", (int)currentCell.Style.Font.FontSize, 11));
            styleBuilder.Append(PropertyToStyle("font-family", currentCell.Style.Font.FontName, "serif"));
            styleBuilder.Append(PropertyToStyle("font-style", currentCell.Style.Font.Italic ? "italic" : "normal", "normal"));
            styleBuilder.Append(PropertyToStyle("font-weight", currentCell.Style.Font.Bold ? "bold" : "normal", "normal"));
            styleBuilder.Append(PropertyToStyle("white-space", currentCell.Style.Alignment.WrapText ? "no-wrap" : "", string.Empty));
            styleBuilder.Append(PropertyToStyle("height", GetRowHeight(range), 0));
            styleBuilder.Append(PropertyToStyle("width", GetColumnWidth(range), 0));
            styleBuilder.Append(PropertyToStyle("transform", currentCell.Style.Alignment.TextRotation, 0));
            return $"<td style=\"{styleBuilder.ToString().TrimEnd()}\" eth-cell=\"{currentCell.Address}\" colspan=\"{range.ColumnCount()}\" rowspan=\"{range.RowCount()}\">";
        }

        /// <summary>
        /// Converts an Excel color to HTML hex format.
        /// </summary>
        /// <param name="cssproperty">The CSS property to set.</param>
        /// <param name="borderStyle">The border style value.</param>
        /// <param name="cellColor">The color of the cell.</param>
        /// <param name="theme">The theme to use for resolving cell styles.</param>
        /// <returns>The HTML representation of the CSS property with the appropriate color.</returns>
        private static string PropertyToStyle(string cssproperty, XLBorderStyleValues borderStyle, XLColor cellColor, IXLTheme theme)
        {
            string width = string.Empty;
            string style = string.Empty;
            switch (borderStyle)
            {
                case XLBorderStyleValues.DashDot:
                    width = "1pt"; //width = "1px";
                    style = "dashed";
                    break;

                case XLBorderStyleValues.DashDotDot:
                    width = "1pt"; //width = "1px";
                    style = "dashed";
                    break;

                case XLBorderStyleValues.Dashed:
                    width = "1pt"; //width = "1px";
                    style = "dashed";
                    break;

                case XLBorderStyleValues.Dotted:
                    width = "1pt"; //width = "1px";
                    style = "dotted";
                    break;

                case XLBorderStyleValues.Double:
                    width = "1pt"; //width = "1px";
                    style = "double";
                    break;

                case XLBorderStyleValues.Hair:
                    width = "1pt"; //width = "1px";
                    style = "solid";
                    break;

                case XLBorderStyleValues.Medium:
                    width = "medium";
                    style = "solid";
                    break;

                case XLBorderStyleValues.MediumDashDot:
                    width = "medium";
                    style = "dashed";
                    break;

                case XLBorderStyleValues.MediumDashDotDot:
                    width = "medium";
                    style = "dotted";
                    break;

                case XLBorderStyleValues.MediumDashed:
                    width = "medium";
                    style = "solid";
                    break;

                case XLBorderStyleValues.SlantDashDot:
                    width = "1pt"; //width = "1px";
                    style = "dashed";
                    break;

                case XLBorderStyleValues.Thick:
                    width = "2.5pt"; //width = "thick";
                    style = "solid";
                    break;

                case XLBorderStyleValues.Thin:
                    width = "thin";
                    style = "solid";
                    break;

                case XLBorderStyleValues.None:
                default:
                    width = "0pt";
                    style = "solid";
                    return string.Empty;
            }

            string color = cellColor.ColorType switch
            {
                XLColorType.Color => ToHex(cellColor.Color),
                XLColorType.Indexed => ToHex(cellColor.Color),
                XLColorType.Theme => ToHex(cellColor.ThemeColor, theme, cellColor.ThemeTint),
                _ => "#000000",
            };
            return PropertyToStyle(cssproperty, $"{style} {width} {color}");
        }

        /// <summary>
        /// Converts an Excel theme color to HTML hex format.
        /// </summary>
        /// <param name="cssproperty">The CSS property to set.</param>
        /// <param name="cellColor">The color of the cell.</param>
        /// <param name="theme">The theme to use for resolving cell styles.</param>
        /// <returns>The HTML representation of the CSS property with the appropriate color.</returns>
        private static string PropertyToStyle(string cssproperty, XLColor cellColor, IXLTheme theme)
        {
            switch (cellColor.ColorType)
            {
                case XLColorType.Color:
                    return PropertyToStyle(cssproperty, ToHex(cellColor.Color));

                case XLColorType.Indexed:
                    return PropertyToStyle(cssproperty, ToHex(cellColor.Color));

                case XLColorType.Theme:
                    return PropertyToStyle(cssproperty, ToHex(cellColor.ThemeColor, theme, cellColor.ThemeTint));

                default:
                    return PropertyToStyle(cssproperty, "#000000");
            }
        }

        /// <summary>
        /// Converts an Excel horizontal alignment value to CSS format.
        /// </summary>
        /// <param name="cssproperty">The CSS property to set.</param>
        /// <param name="value">The horizontal alignment value.</param>
        /// <returns>The HTML representation of the CSS property with the appropriate alignment.</returns>
        private static string PropertyToStyle(string cssproperty, XLAlignmentHorizontalValues value)
        {
            var temp = value.ToString() ?? "general";
            return PropertyToStyle(cssproperty, temp);
        }

        /// <summary>
        /// Converts an Excel vertical alignment value to CSS format.
        /// </summary>
        /// <param name="cssproperty">The CSS property to set.</param>
        /// <param name="value">The vertical alignment value.</param>
        /// <returns>The HTML representation of the CSS property with the appropriate alignment.</returns>
        private static string PropertyToStyle(string cssproperty, XLAlignmentVerticalValues value)
        {
            var temp = value switch
            {
                XLAlignmentVerticalValues.Bottom => "bottom",
                XLAlignmentVerticalValues.Center => "middle",
                XLAlignmentVerticalValues.Top => "top",
                _ => "bottom",
            };
            return PropertyToStyle(cssproperty, temp);
        }

        /// <summary>
        /// Converts an integer value to CSS format, considering a default value.
        /// </summary>
        /// <param name="cssproperty">The CSS property name to set.</param>
        /// <param name="value">The CSS property value</param>
        /// <param name="defaultValue">The CSS property default value</param>
        /// <returns>The HTML representation of the CSS property with the appropriate value.</returns>
        private static string PropertyToStyle(string cssproperty, int value, int defaultValue)
        {
            if (value == defaultValue)
            {
                return string.Empty;
            }

            return PropertyToStyle(cssproperty, value.ToString());
        }

        /// <summary>
        /// Converts an integer value to CSS format, considering a default value.
        /// </summary>
        /// <param name="cssproperty">The CSS property name to set.</param>
        /// <param name="value">The CSS property value</param>
        /// <param name="defaultValue">The CSS property default value</param>
        /// <returns>The HTML representation of the CSS property with the appropriate value.</returns>
        private static string PropertyToStyle(string cssproperty, double value, double defaultValue)
        {
            if (Math.Abs(value - defaultValue) < 0.001)
            {
                return string.Empty;
            }

            return PropertyToStyle(cssproperty, value.ToString("F2"));
        }

        /// <summary>
        /// Converts an integer value to CSS format, considering a default value.
        /// </summary>
        /// <param name="cssproperty">The CSS property name to set.</param>
        /// <param name="value">The CSS property value</param>
        /// <param name="defaultValue">The CSS property default value</param>
        /// <returns>The HTML representation of the CSS property with the appropriate value.</returns>
        private static string PropertyToStyle(string cssproperty, string value, string defaultValue)
        {
            if (value == defaultValue)
            {
                return string.Empty;
            }

            return PropertyToStyle(cssproperty, value);
        }

        /// <summary>
        /// Converts an integer value to CSS format, considering a default value.
        /// </summary>
        /// <param name="cssproperty">The CSS property name to set.</param>
        /// <param name="value">The CSS property value</param>
        /// <returns>The HTML representation of the CSS property with the appropriate value.</returns>
        private static string PropertyToStyle(string cssproperty, string value)
        {
            if (cssproperty.Contains("size") ||
                cssproperty.Contains("width") ||
                cssproperty.Contains("height"))
            {
                var temp = value.Replace(",", ".");
                return $"{cssproperty}:{temp}pt; ";
            }
            if (cssproperty.Contains("transform"))
            {
                return $"{cssproperty}:rotate(-{value}deg); ";
            }

            return $"{cssproperty}:{value}; ";
        }

        /// <summary>
        /// Converts an RGB color to HTML hex format.
        /// </summary>
        /// <param name="c">The RGB color to convert.</param>
        /// <returns>The HTML hex representation of the RGB color.</returns>
        private static string ToHex(System.Drawing.Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        /// <summary>
        /// Converts an Excel theme color to HTML hex format.
        /// </summary>
        /// <param name="themeColor">The Excel theme color to convert.</param>
        /// <param name="theme">The Excel theme to use for resolving the color.</param>
        /// <returns>The HTML hex representation of the RGB color.</returns>
        private static string ToHex(XLThemeColor themeColor, IXLTheme theme)
        {
            var color = theme.ResolveThemeColor(themeColor).Color;
            return ToHex(color);
        }

        /// <summary>
        /// Converts an Excel theme color to HTML hex format.
        /// </summary>
        /// <param name="themeColor">The Excel theme color to convert.</param>
        /// <param name="theme">The Excel theme to use for resolving the color.</param>
        /// <param name="tint">The tint value to apply to the theme color.</param>
        /// <returns>The HTML hex representation of the RGB color.</returns>
        private static string ToHex(XLThemeColor themeColor, IXLTheme theme, double tint)
        {
            var color = theme.ResolveThemeColor(themeColor).Color;
            color = ApplyTint(color, tint);
            return ToHex(color);
        }

        //static System.Drawing.Color ApplyTint(System.Drawing.Color originalColor, double tintFactor)
        //{
        //    int red = (int)(originalColor.R * (1 + tintFactor));
        //    int green = (int)(originalColor.G * (1 + tintFactor));
        //    int blue = (int)(originalColor.B * (1 + tintFactor));

        //    // 범위를 초과하는 경우 최대값으로 설정
        //    red = Math.Min(255, red);
        //    green = Math.Min(255, green);
        //    blue = Math.Min(255, blue);

        //    return System.Drawing.Color.FromArgb(originalColor.A, red, green, blue);
        //}

        /// <summary>
        /// Converts an RGB color to HTML RGB format.
        /// </summary>
        /// <param name="c">The RGB color to convert.</param>
        /// <returns>The HTML RGB representation of the RGB color.</returns>
        private static string ToRgb(System.Drawing.Color c) => $"RGB({c.R},{c.G},{c.B})";

        /// <summary>
        /// Applies a tint to an RGB color.
        /// </summary>
        /// <param name="color">The RGB color to which the tint will be applied.</param>
        /// <param name="tint">The tint value to apply.</param>
        /// <returns>The RGB color after applying the tint.</returns>
        private static System.Drawing.Color ApplyTint(System.Drawing.Color color, double tint)
        {
            if (tint < -1.0) tint = -1.0;
            if (tint > 1.0) tint = 1.0;

            System.Drawing.Color colorRgb = color;
            double hue = colorRgb.GetHue();
            double saturation = colorRgb.GetSaturation();
            double lum = colorRgb.GetBrightness();
            if (tint < 0)
            {
                lum = lum * (1.0 + tint);
            }
            else
            {
                lum = lum * (1.0 - tint) + (1.0 - 1.0 * (1.0 - tint));
            }

            return ToColor(hue, saturation, lum);
        }

        /// <summary>
        /// Converts hue, saturation, and luminance values to an RGB color.
        /// </summary>
        /// <param name="hue">The hue value of the color.</param>
        /// <param name="saturation">The saturation value of the color.</param>
        /// <param name="luminance">The luminance value of the color.</param>
        /// <returns>The RGB color calculated from the provided hue, saturation, and luminance values.</returns>
        private static System.Drawing.Color ToColor(double hue, double saturation, double luminance)
        {
            double chroma = (1.0 - Math.Abs(2.0 * luminance - 1.0)) * saturation;
            double fHue = hue / 60.0;
            double fHueMod2 = fHue;
            while (fHueMod2 >= 2.0) fHueMod2 -= 2.0;
            double fTemp = chroma * (1.0 - Math.Abs(fHueMod2 - 1.0));

            double fRed, fGreen, fBlue;
            if (fHue < 1.0)
            {
                fRed = chroma;
                fGreen = fTemp;
                fBlue = 0;
            }
            else if (fHue < 2.0)
            {
                fRed = fTemp;
                fGreen = chroma;
                fBlue = 0;
            }
            else if (fHue < 3.0)
            {
                fRed = 0;
                fGreen = chroma;
                fBlue = fTemp;
            }
            else if (fHue < 4.0)
            {
                fRed = 0;
                fGreen = fTemp;
                fBlue = chroma;
            }
            else if (fHue < 5.0)
            {
                fRed = fTemp;
                fGreen = 0;
                fBlue = chroma;
            }
            else if (fHue < 6.0)
            {
                fRed = chroma;
                fGreen = 0;
                fBlue = fTemp;
            }
            else
            {
                fRed = 0;
                fGreen = 0;
                fBlue = 0;
            }

            double fMin = luminance - 0.5 * chroma;
            fRed += fMin;
            fGreen += fMin;
            fBlue += fMin;

            fRed *= 255.0;
            fGreen *= 255.0;
            fBlue *= 255.0;

            var red = System.Convert.ToInt32(Math.Truncate(fRed));
            var green = System.Convert.ToInt32(Math.Truncate(fGreen));
            var blue = System.Convert.ToInt32(Math.Truncate(fBlue));

            red = Math.Min(255, Math.Max(red, 0));
            green = Math.Min(255, Math.Max(green, 0));
            blue = Math.Min(255, Math.Max(blue, 0));

            return System.Drawing.Color.FromArgb(red, green, blue);
        }

        /// <summary>
        /// Gets the width of a column in points based on the width of the provided cell.
        /// </summary>
        /// <param name="cell">The cell whose column's width is to be determined.</param>
        /// <returns>The width of the column in points.</returns>
        private static double GetColumnWidth(IXLCell cell)
        {
            return Math.Round(cell.WorksheetColumn().Width * 5, 2);
        }

        /// <summary>
        /// Gets the total width of columns within a range in points.
        /// </summary>
        /// <param name="range">The range of cells whose columns' widths are to be summed.</param>
        /// <returns>The total width of columns within the range in points.</returns>
        private static double GetColumnWidth(IXLRange range)
        {
            double totalWidth = 0;
            foreach (var column in range.Columns())
            {
                totalWidth += Math.Round(column.WorksheetColumn().Width * 5, 2);
            }
            return totalWidth;
        }

        /// <summary>
        /// Gets the height of a row in points.
        /// </summary>
        /// <param name="cell">The cell in the row whose height is to be retrieved.</param>
        /// <returns>The height of the row containing the specified cell in points.</returns>
        private static double GetRowHeight(IXLCell cell)
        {
            return Math.Round(cell.WorksheetRow().Height, 2);
        }

        /// <summary>
        /// Gets the total height of rows in a range in points.
        /// </summary>
        /// <param name="range">The range of cells whose row heights are to be summed.</param>
        /// <returns>The total height of rows in the specified range in points.</returns>
        private static double GetRowHeight(IXLRange range)
        {
            double totalHeight = 0;
            foreach (var row in range.Rows())
            {
                totalHeight += Math.Round(row.WorksheetRow().Height, 2);
            }
            return totalHeight;
        }

        /// <summary>
        /// Encodes special characters in HTML.
        /// </summary>
        /// <param name="input">The input string to encode.</param>
        /// <returns>The HTML-encoded string.</returns>
        private static string EncodeCellValue(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return "&nbsp;";
            }

            // HTML encoding
            input = System.Net.WebUtility.HtmlEncode(input);
            // Line break conversion
            input = input.Replace("\n", "<br>\n");
            // Convert leading spaces to HTML entities
            input = Regex.Replace(input, @"^(\s+)",
                (m) =>
                {
                    StringBuilder sb = new StringBuilder();
                    for (var i = 0; i < m.Length; i++)
                    {
                        sb.Append("&nbsp;");
                    }

                    return sb.ToString();
                }, RegexOptions.Multiline);
            return input;
        }
    }
}