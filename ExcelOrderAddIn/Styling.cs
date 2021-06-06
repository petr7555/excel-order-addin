using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public static class Styling
    {
        private static readonly Color Grey = ColorTranslator.FromHtml("#C0C0C0");
        private static readonly Color Salmon = ColorTranslator.FromHtml("#FCE4D6");
        private static readonly Color Yellow = ColorTranslator.FromHtml("#FFF2CC");
        private static readonly Color Red = ColorTranslator.FromHtml("#FF0000");

        public enum Style
        {
            Calculation, // Exists in Excel by default
            Input, // Exists in Excel by default
            Header,
            SalmonBold,
            Yellow,
            BoldText,
            RedBoldText,
            RedBoldHeaderText,
        }

        public static void Apply(Excel.Range range, Style style)
        {
            switch (style)
            {
                case Style.Calculation:
                    ApplyCalculation(range);
                    break;
                case Style.Input:
                    ApplyInput(range);
                    break;
                case Style.Header:
                    ApplyHeader(range);
                    break;
                case Style.SalmonBold:
                    ApplySalmonBold(range);
                    break;
                case Style.Yellow:
                    ApplyYellow(range);
                    break;
                case Style.BoldText:
                    ApplyBoldText(range);
                    break;
                case Style.RedBoldText:
                    ApplyRedBoldText(range);
                    break;
                case Style.RedBoldHeaderText:
                    ApplyRedBoldHeaderText(range);
                    break;
                default:
                    throw new NotImplementedException($"The style {style} is not implemented.");
            }
        }

        private static void ApplyCalculation(Excel.Range range)
        {
            range.Style = "Calculation";
        }

        private static void ApplyInput(Excel.Range range)
        {
            range.Style = "Input";
        }


        private static void ApplyHeader(Excel.Range range)
        {
            const string styleName = "Header_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(Grey)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplySalmonBold(Excel.Range range)
        {
            const string styleName = "SalmonBold_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(Salmon)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplyYellow(Excel.Range range)
        {
            const string styleName = "Yellow_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(Yellow);
            range.Style = styleName;
        }

        private static void ApplyBoldText(Excel.Range range)
        {
            const string styleName = "BoldText_addin";
            new StyleBuilder(styleName)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplyRedBoldText(Excel.Range range)
        {
            const string styleName = "RedBoldText_addin";
            new StyleBuilder(styleName)
                .WithTextColor(Red)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplyRedBoldHeaderText(Excel.Range range)
        {
            const string styleName = "RedBoldHeaderText_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(Grey)
                .WithTextColor(Red)
                .WithBold();
            range.Style = styleName;
        }


        public class StyleBuilder
        {
            Excel.Style Style = null;

            public StyleBuilder(string styleName)
            {
                try
                {
                    Style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add(styleName);
                }
                catch (COMException e) when (e.Message ==
                                             "Add method of Styles class failed")
                {
                }
            }

            public StyleBuilder WithFontName(string fontName)
            {
                if (Style != null)
                {
                    Style.Font.Name = fontName;
                }

                return this;
            }

            public StyleBuilder WithFontSize(int fontSize)
            {
                if (Style != null)
                {
                    Style.Font.Size = fontSize;
                }

                return this;
            }

            public StyleBuilder WithBold()
            {
                if (Style != null)
                {
                    Style.Font.Bold = true;
                }

                return this;
            }

            public StyleBuilder WithTextColor(Color textColor)
            {
                if (Style != null)
                {
                    Style.Font.Color = textColor;
                }

                return this;
            }

            public StyleBuilder WithBackgroundColor(Color backgroundColor)
            {
                if (Style == null) return this;
                Style.Interior.Color = backgroundColor;
                Style.Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                return this;
            }
        }
    }
}
