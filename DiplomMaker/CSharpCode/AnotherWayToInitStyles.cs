using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public partial class MakeDoc
    {
        // Пока просто заготовка
        private Word.Style GetNewParagraphStyle
        (
            string styleName,
            string fontName = null,
            int? fontSize = null,
            bool? isFontBold = null,
            bool? isFontItalic = null,
            bool? isFontStrikeThrough = null,
            Word.WdUnderline? fontUnderline = null,
            Word.WdColor? fontUnderlineColor = null,
            Word.WdColor? fontColor = null,
            float? leftIndent = null,
            float? rightIndent = null,
            float? firstLineIndent = null,
            float? spaceAfter = null,
            float? spaceBefore = null,
            Word.WdLineSpacing? lineSpacingRule = null,
            float? lineSpacing= null,
            Word.WdParagraphAlignment? alignment = null,
            bool? keepWithNext = null,
            bool? keepTogether = null,
            Word.WdOutlineLevel? wdOutlineLevel = null,
            bool? allCaps = null
        )
        {
            var style = WordDoc.Styles.Add(styleName, Word.WdStyleType.wdStyleTypeParagraph);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            style.AutomaticallyUpdate = false;

            // font
            if (fontName != null)  font.Name = fontName; 
            if (fontSize != null)  font.Size = fontSize.Value; 
            if (fontColor != null)  font.Color = fontColor.Value; 
            if (fontUnderline != null)  font.Underline = fontUnderline.Value; 
            if (fontUnderlineColor != null)  font.UnderlineColor = fontUnderlineColor.Value; 
            if (isFontBold != null) font.Bold = isFontBold.Value ? True : False;
            if (isFontItalic != null) font.Italic = isFontItalic.Value ? True : False;
            if (allCaps != null) font.AllCaps = allCaps.Value ? True : False;
            if (isFontStrikeThrough != null) font.StrikeThrough = isFontStrikeThrough.Value ? True : False;

            // paragraphFormat
            if (leftIndent!=null) paragraphFormat.LeftIndent = WordApp.CentimetersToPoints(leftIndent.Value);
            if (rightIndent != null) paragraphFormat.RightIndent = WordApp.CentimetersToPoints(rightIndent.Value);
            if (firstLineIndent != null) paragraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(firstLineIndent.Value);
            if (spaceAfter != null) paragraphFormat.SpaceAfter = spaceAfter.Value;
            if (spaceBefore != null) paragraphFormat.SpaceBefore = spaceBefore.Value;
            if (lineSpacingRule != null) paragraphFormat.LineSpacingRule = lineSpacingRule.Value;
            if (lineSpacingRule == Word.WdLineSpacing.wdLineSpaceMultiple && lineSpacing != null)
                paragraphFormat.LineSpacing = WordApp.LinesToPoints(lineSpacing.Value);
            if (alignment != null) paragraphFormat.Alignment = alignment.Value;
            if (keepWithNext != null) paragraphFormat.KeepWithNext = keepWithNext.Value ? True : False;
            if (keepTogether != null) paragraphFormat.KeepTogether = keepTogether.Value ? True : False;
            if (wdOutlineLevel != null) paragraphFormat.OutlineLevel = wdOutlineLevel.Value;
            paragraphFormat.TabStops.ClearAll();

            return style;
        }

    }
}
