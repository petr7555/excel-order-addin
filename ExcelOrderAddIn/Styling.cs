using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public static class Styling
    {

        //private static readonly Color PURPLE = ColorTranslator.FromHtml("#3F3F76");
        //private static readonly Color ORANGE = ColorTranslator.FromHtml("#FFCC99");
        private static readonly Color GREY = ColorTranslator.FromHtml("#C0C0C0");
        private static readonly Color BLACK = ColorTranslator.FromHtml("#000000");
        private static readonly Color WHITE = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color SALMON = ColorTranslator.FromHtml("#FCE4D6");
        private static readonly Color YELLOW = ColorTranslator.FromHtml("#FFF2CC");
        private static readonly Color RED = ColorTranslator.FromHtml("#FF0000");


        public enum Style
        {
            CALCULATION, // Exists
            INPUT, // Exists
            HEADER,
            SALMON,
            YELLOW,
            BOLD_TEXT,
            RED_TEXT,
        }

        public static void Apply(Excel.Range range, Style style)
        {
            switch (style)
            {
                case Style.CALCULATION:
                    ApplyCalculation(range);
                    break;
                case Style.INPUT:
                    ApplyInput(range);
                    break;
                case Style.HEADER:
                    ApplyHeader(range);
                    break;
                case Style.SALMON:
                    ApplySalmon(range);
                    break;
                case Style.YELLOW:
                    ApplyYellow(range);
                    break;
                case Style.BOLD_TEXT:
                    ApplyBoldText(range);
                    break;
                case Style.RED_TEXT:
                    ApplyRedText(range);
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
            var styleName = "Header_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(GREY)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplySalmon(Excel.Range range)
        {
            var styleName = "Salmon_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(SALMON);
            range.Style = styleName;
        }

        private static void ApplyYellow(Excel.Range range)
        {
            var styleName = "Yellow_addin";
            new StyleBuilder(styleName)
                .WithBackgroundColor(YELLOW);
            range.Style = styleName;
        }

        private static void ApplyBoldText(Excel.Range range)
        {
            var styleName = "BoldText_addin";
            new StyleBuilder(styleName)
                .WithBold();
            range.Style = styleName;
        }

        private static void ApplyRedText(Excel.Range range)
        {
            var styleName = "RedText_addin";
            new StyleBuilder(styleName)
                .WithTextColor(RED);
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
                catch (System.Runtime.InteropServices.COMException e) when (e.Message == "Add method of Styles class failed")
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
                if (Style != null)
                {
                    Style.Interior.Color = backgroundColor;
                    Style.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                }
                return this;
            }
        }
    }
}
