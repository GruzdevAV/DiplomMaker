using Microsoft.VisualBasic;
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
        // Для параграфов
        public static string DefaultFontName_forStyle { get; set; } = "Times New Roman";
        public static int? DefaultFontSize { get; set; } = 14;
        public static bool? DefaultBold { get; set; } = null;
        public static bool? DefaultItalic { get; set; } = null;
        public static bool? DefaultStrikeThrough { get; set; } = null;
        public static Word.WdUnderline? DefaultUnderline { get; set; } = null;
        public static Word.WdColor? DefaultUnderlineColor { get; set; } = null;
        public static Word.WdColor? DefaultFontColor { get; set; } = Word.WdColor.wdColorBlack;
        public static float? DefaultLeftIndent { get; set; } = null;
        public static float? DefaultRightIndent { get; set; } = null;
        public static float? DefaultFirstLineIndent { get; set; } = null;
        public static float? DefaultSpaceAfter { get; set; } = 0;
        public static float? DefaultSpaceBefore { get; set; } = 0;
        public static Word.WdLineSpacing? DefaultLineSpacingRule { get; set; } = Word.WdLineSpacing.wdLineSpace1pt5;
        public static float? DefaultLineSpacing { get; set; } = null;
        public static Word.WdParagraphAlignment? DefaultParagraphAlignment { get; set; } = Word.WdParagraphAlignment.wdAlignParagraphJustify;
        public static bool? DefaultKeepWithNext { get; set; } = null;
        public static bool? defaultKeepTogether { get; set; } = null;
        public static Word.WdOutlineLevel? DefaultOutlineLevel { get; set; } = null;
        public static bool? DefaultAllCaps { get; set; } = null;
        // Для списков
        public static string DefaultFontName_ForListLevel { get; set; } = null;
        public static string DefaultNumberFormat { get; set; } = null;
        public static Word.WdTrailingCharacter? DefaultTrailingCharacter { get; set; } = Word.WdTrailingCharacter.wdTrailingTab;
        public static Word.WdListNumberStyle? DefaultNumberStyle { get; set; } = null;
        public static float? DefaultNumberPosition { get; set; } =1.25f;
        public static Word.WdListLevelAlignment? DefaultListLevelAlignment { get; set; } = null;
        public static float? DefaultTextPosition { get; set; } = 0;
        public static float? DefaultTabPosition { get; set; } = 2;
        public static int? DefaultResetOnHigher { get; set; } = null;
        public static int? DefaultStartAt { get; set; } = null;
        public static Word.Style DefaultLinkedStyle { get; set; } = null;

        /// <summary>
        /// Получить новый стиль для абзаца
        /// </summary>
        /// <param name="styleName">Имя создаваемого стиля</param>
        /// <param name="fontName">Шрифт</param>
        /// <param name="fontSize">Кегель шрифта</param>
        /// <param name="bold">Жирный ли</param>
        /// <param name="italic">Наклонный ли</param>
        /// <param name="strikeThrough">Перечёркнутый ли</param>
        /// <param name="underline">Подчёркивание</param>
        /// <param name="underlineColor">Цвет подчёркивания</param>
        /// <param name="fontColor">Цвет шрифта</param>
        /// <param name="leftIndent">Отступ слева</param>
        /// <param name="rightIndent">Отступ справа</param>
        /// <param name="firstLineIndent">Отступ первой строки</param>
        /// <param name="spaceAfter">Отступ перед текстом</param>
        /// <param name="spaceBefore">Отступ после текста</param>
        /// <param name="lineSpacingRule">Правило междустрочного интервала</param>
        /// <param name="lineSpacing">Значение междустрочного интервала</param>
        /// <param name="alignment">Выравнивание</param>
        /// <param name="keepWithNext">Не отрывать от следующего</param>
        /// <param name="keepTogether">Не разрывать абзац</param>
        /// <param name="outlineLevel">Уровень заголовка (10 - обычный текст)</param>
        /// <param name="allCaps">Все заглавные</param>
        /// <returns>Стиль абзаца</returns>
        private Word.Style GetNewParagraphStyle
                (
                    string styleName,
                    string fontName = null,
                    int? fontSize = null,
                    bool? bold = null,
                    bool? italic = null,
                    bool? strikeThrough = null,
                    Word.WdUnderline? underline = null,
                    Word.WdColor? underlineColor = null,
                    Word.WdColor? fontColor = null,
                    float? leftIndent = null,
                    float? rightIndent = null,
                    float? firstLineIndent = null,
                    float? spaceAfter = null,
                    float? spaceBefore = null,
                    Word.WdLineSpacing? lineSpacingRule = null,
                    float? lineSpacing = null,
                    Word.WdParagraphAlignment? alignment = null,
                    bool? keepWithNext = null,
                    bool? keepTogether = null,
                    Word.WdOutlineLevel? outlineLevel = null,
                    bool? allCaps = null
                )
        {

            fontName = fontName ?? DefaultFontName_forStyle;
            fontSize = fontSize ?? DefaultFontSize;
            bold = bold ?? DefaultBold;
            italic = italic ?? DefaultItalic;
            strikeThrough = strikeThrough ?? DefaultStrikeThrough;
            underline = underline ?? DefaultUnderline;
            underlineColor = underlineColor ?? DefaultUnderlineColor;
            fontColor = fontColor ?? DefaultFontColor;
            leftIndent = leftIndent ?? DefaultLeftIndent;
            rightIndent = rightIndent ?? DefaultRightIndent;
            firstLineIndent = firstLineIndent ?? DefaultFirstLineIndent;
            spaceAfter = spaceAfter ?? DefaultSpaceAfter;
            spaceBefore = spaceBefore ?? DefaultSpaceBefore;
            lineSpacingRule = lineSpacingRule ?? DefaultLineSpacingRule;
            lineSpacing = lineSpacing ?? DefaultLineSpacing;
            alignment = alignment ?? DefaultParagraphAlignment;
            keepWithNext = keepWithNext ?? DefaultKeepWithNext;
            keepTogether = keepTogether ?? defaultKeepTogether;
            outlineLevel = outlineLevel ?? DefaultOutlineLevel;
            allCaps = allCaps ?? DefaultAllCaps;

            var style = WordDoc.Styles.Add(styleName, Word.WdStyleType.wdStyleTypeParagraph);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            style.AutomaticallyUpdate = false;

            // font
            if (fontName != null) font.Name = fontName;
            if (fontSize != null) font.Size = fontSize.Value;
            if (fontColor != null) font.Color = fontColor.Value;
            if (underline != null) font.Underline = underline.Value;
            if (underlineColor != null) font.UnderlineColor = underlineColor.Value;
            if (bold != null) font.Bold = bold.Value ? True : False;
            if (italic != null) font.Italic = italic.Value ? True : False;
            if (allCaps != null) font.AllCaps = allCaps.Value ? True : False;
            if (strikeThrough != null) font.StrikeThrough = strikeThrough.Value ? True : False;

            // paragraphFormat
            if (leftIndent != null) paragraphFormat.LeftIndent = WordApp.CentimetersToPoints(leftIndent.Value);
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
            if (outlineLevel != null) paragraphFormat.OutlineLevel = outlineLevel.Value;
            paragraphFormat.TabStops.ClearAll();

            return style;
        }
        private Word.Style GetStyleMainHeading(string name = "MainHeading")
        {
            var style = GetNewParagraphStyle
                (
                name,
                bold: true,
                allCaps: true,
                leftIndent: 1.25f,
                spaceAfter: 18,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphCenter,
                keepTogether: true,
                keepWithNext: true,
                outlineLevel: Word.WdOutlineLevel.wdOutlineLevel1
                );
            return style;
        }
        private Word.Style GetStyleRegularText(string name = "RegularText")
        {
            var style = GetNewParagraphStyle
                (
                name,
                firstLineIndent: 1.25f
                );
            return style;
        }
        private Word.Style GetStyleFormulaParagraph(string name = "FormulaParagraph")
        {
            var style = GetNewParagraphStyle
                (
                name,
                spaceAfter: 12,
                spaceBefore: 12,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphLeft
                );
            var paragraphFormat = style.ParagraphFormat;
            paragraphFormat.TabStops.Add(WordApp.CentimetersToPoints(8.25f), Word.WdTabAlignment.wdAlignTabCenter, Word.WdTabLeader.wdTabLeaderSpaces);
            paragraphFormat.TabStops.Add(WordApp.CentimetersToPoints(15.5f), Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderSpaces);
            return style;
        }
        private Word.Style GetStyleImageParagraph(string name = "ImageParagraph")
        {
            var style = GetNewParagraphStyle
                (
                name,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphCenter,
                keepWithNext: true,
                keepTogether: true
                );
            return style;
        }
        private Word.Style GetStyleImageNumberParagraph(string name = "ImageNumberParagraph")
        {
            var style = GetNewParagraphStyle
                (
                name,
                spaceAfter: 18,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphCenter
                );
            return style;
        }
        private Word.Style GetStyleTableNumberParagraph(string name = "TableNumberParagraph")
        {
            var style = GetNewParagraphStyle
                (
                name,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphLeft
                );
            return style;
        }
        private Word.Style GetStyleListHeading(string name, int level, bool afterText)
        {
            var style = GetNewParagraphStyle
                (
                name,
                spaceBefore: (afterText ? 24 : (level == 1 ? 0 : 12)),
                spaceAfter: 12,
                alignment: Word.WdParagraphAlignment.wdAlignParagraphLeft,
                keepWithNext: true,
                keepTogether: true,
                firstLineIndent: 1.25f,
                outlineLevel: (Word.WdOutlineLevel)level,
                bold: true
                );
            return style;
        }
        private Word.Style GetStyleAfterTable(string name, bool beforeText)
        {
            var style = GetNewParagraphStyle
                (
                name,
                spaceAfter: (beforeText ? 30 : 36)
                );
            return style;
        }
        private Word.ListLevel GetNewListLevel
            (
            
            Word.WdListGalleryType listGalleryType,
            int level = 1,
            string fontName = null,
            string numberFormat = null,
            Word.WdTrailingCharacter? trailingCharacter = null,
            Word.WdListNumberStyle? numberStyle = null,
            float? numberPosition = null,
            Word.WdListLevelAlignment? alignment = null,
            float? textPosition = null,
            float? tabPosition = null,
            int? resetOnHigher = null,
            int? startAt = null,
            Word.Style linkedStyle = null
            )
        {
            fontName = fontName ?? DefaultFontName_ForListLevel;
            numberFormat = numberFormat ?? DefaultNumberFormat;
            trailingCharacter = trailingCharacter ?? DefaultTrailingCharacter;
            numberStyle = numberStyle ?? DefaultNumberStyle;
            numberPosition = numberPosition ?? DefaultNumberPosition;
            alignment = alignment ?? DefaultListLevelAlignment;
            textPosition = textPosition ?? DefaultTextPosition;
            tabPosition = tabPosition ?? DefaultTabPosition;
            resetOnHigher = resetOnHigher ?? DefaultResetOnHigher;
            startAt = startAt ?? DefaultStartAt;
            linkedStyle = linkedStyle ?? DefaultLinkedStyle;


            var listGallery = WordApp.ListGalleries[listGalleryType].ListTemplates[1].ListLevels[level];
            if (fontName != null) listGallery.Font.Name = fontName;
            if (numberFormat != null) listGallery.NumberFormat = numberFormat;
            if (trailingCharacter != null) listGallery.TrailingCharacter = trailingCharacter.Value;
            if (numberStyle != null) listGallery.NumberStyle = numberStyle.Value;
            if (numberPosition != null) listGallery.NumberPosition = WordApp.CentimetersToPoints(numberPosition.Value);
            if (alignment != null) listGallery.Alignment = alignment.Value;
            if (textPosition != null) listGallery.TextPosition = WordApp.CentimetersToPoints(textPosition.Value);
            if (tabPosition != null) listGallery.TabPosition = WordApp.CentimetersToPoints(tabPosition.Value);
            if (resetOnHigher != null) listGallery.ResetOnHigher = resetOnHigher.Value;
            if (startAt != null) listGallery.StartAt = startAt.Value;
            if (linkedStyle != null) listGallery.LinkedStyle = linkedStyle.NameLocal;

            return listGallery;
        }
        private Word.Style GetStyleDashList(string name = "DashList")
        {
            var lg = GetNewListLevel
                (
                listGalleryType: Word.WdListGalleryType.wdBulletGallery,
                fontName:"Symbol",
                numberStyle: Word.WdListNumberStyle.wdListNumberStyleBullet,
                numberFormat: Strings.ChrW(61630).ToString(),
                alignment: Word.WdListLevelAlignment.wdListLevelAlignLeft,
                resetOnHigher:0,
                startAt:1
                );
            var style = GetNewParagraphStyle(name);

            style.LinkToListTemplate(
            WordApp.ListGalleries[Word.WdListGalleryType.wdBulletGallery]
                .ListTemplates[1], 1);
            style.set_BaseStyle(RegularText.NameLocal);
            lg.LinkedStyle = style.NameLocal;
            return style;
        }
        private Word.Style GetStyleForListHeadings(string name = "ListHeadingsVersion1")
        {
            var style = GetNewParagraphStyle
                (
                name,
                bold:true
                );
            
            var fonts = new[] { ListHeading1, ListHeading2AfterHead, ListHeading3AfterHead, ListHeading4AfterHead };
            var stringList = new List<string>();
            for (int level = 1; level <= fonts.Length; level++)
            {
                stringList.Add($"%{level}");
                var listLevelN = GetNewListLevel
                    (
                    Word.WdListGalleryType.wdOutlineNumberGallery,
                    level,
                    numberFormat: string.Join(".", stringList),
                    numberStyle: Word.WdListNumberStyle.wdListNumberStyleArabic,
                    resetOnHigher: level - 1,
                    startAt: 1,
                    alignment: Word.WdListLevelAlignment.wdListLevelAlignLeft,
                    linkedStyle: fonts[level - 1]
                    );
            }
            style.LinkToListTemplate(WordApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1], 1);
            return style;
        }
    }
}
