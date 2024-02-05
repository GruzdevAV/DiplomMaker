using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media.TextFormatting;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public partial class MakeDoc
    {
        private Word.Style GetStyleMainHeading(string name = "MainHeading")
        {
            var style = GetDefaultStyle(name);
            //var style = WordDoc.Styles.Add(name, Word.WdStyleType.wdStyleTypeParagraph);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            font.Bold = True;
            font.AllCaps = True;
            paragraphFormat.LeftIndent = WordApp.CentimetersToPoints(1.25f);
            paragraphFormat.SpaceAfter = 18;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraphFormat.KeepWithNext = True;
            paragraphFormat.KeepTogether = True;
            paragraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
            return style;
        }
        private Word.Style GetStyleRegularText(string name = "RegularText")
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(1.25f);
            return style;
        }
        private Word.Style GetStyleFormulaParagraph(string name = "FormulaParagraph")
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.SpaceBefore = 12;
            paragraphFormat.SpaceAfter = 12;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraphFormat.TabStops.Add(WordApp.CentimetersToPoints(8.25f), Word.WdTabAlignment.wdAlignTabCenter, Word.WdTabLeader.wdTabLeaderSpaces);
            paragraphFormat.TabStops.Add(WordApp.CentimetersToPoints(15.5f), Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderSpaces);
            return style;
        }
        private Word.Style GetStyleImageParagraph(string name = "ImageParagraph")
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraphFormat.KeepWithNext = True;
            paragraphFormat.KeepTogether = True;
            return style;
        }
        private Word.Style GetStyleImageNumberParagraph(string name = "ImageNumberParagraph")
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.SpaceAfter = 18;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            return style;
        }
        private Word.Style GetStyleTableNumberParagraph(string name = "TableNumberParagraph")
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            return style;
        }
        private Word.Style GetStyleListHeading(string name, int level, bool afterText)
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.SpaceBefore = (afterText ? 24 : (level == 1 ? 0 : 12));
            paragraphFormat.SpaceAfter = 12;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraphFormat.KeepWithNext = True;
            paragraphFormat.KeepTogether = True;
            paragraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(1.25f);
            paragraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
            font.Bold = True;
            return style;
        }
        private Word.Style GetStyleAfterTable(string name, bool beforeText)
        {
            var style = GetDefaultStyle(name);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            paragraphFormat.SpaceAfter = (beforeText ? 30 : 36);
            return style;
        }
        private Word.Style GetStyleDashList(string name = "DashList")
        {
            var lg = WordApp.ListGalleries[Word.WdListGalleryType.wdBulletGallery].ListTemplates[1].ListLevels[1];

            var font = lg.Font;
            font.Name = "Symbol";
            lg.NumberFormat = Strings.ChrW(61630).ToString();
            lg.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lg.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleBullet;
            lg.NumberPosition = WordApp.CentimetersToPoints(1.25f);
            lg.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lg.TextPosition = WordApp.CentimetersToPoints(0);
            lg.TabPosition = WordApp.CentimetersToPoints(2);
            lg.ResetOnHigher = 0;
            lg.StartAt = 1;
            //lg.LinkedStyle = "RegularText";
            var style = WordDoc.Styles.Add(name);
            style.LinkToListTemplate(
        WordApp.ListGalleries[Word.WdListGalleryType.wdBulletGallery]
        .ListTemplates[1], 1);
            style.set_BaseStyle(RegularText.NameLocal);
            lg.LinkedStyle = style.NameLocal;
            return style;
        }
        private Word.Style GetStyleForListHeadings(string name = "ListHeadingsVersion1")
        {
            var style = WordDoc.Styles.Add(name);
            var font = style.Font;
            font.Name = "Times New Roman";
            font.Size = 14;
            font.Bold = True;
            font.Italic = False;
            var fonts = new[] { ListHeading1.NameLocal, ListHeading2AfterHead.NameLocal, ListHeading3AfterHead.NameLocal, ListHeading4AfterHead.NameLocal };
            var stringList = new List<string>();
            for (int n = 1; n <= fonts.Length; n++)
            {
                var listLevelN = WordApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1].ListLevels[n];
                stringList.Add($"%{n}");
                listLevelN.NumberFormat = string.Join(".", stringList);
                listLevelN.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
                listLevelN.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                listLevelN.NumberPosition = WordApp.CentimetersToPoints(1.25f);
                listLevelN.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                listLevelN.TextPosition = WordApp.CentimetersToPoints(0);
                listLevelN.TabPosition = WordApp.CentimetersToPoints(2);
                listLevelN.ResetOnHigher = n - 1;
                listLevelN.StartAt = 1;
                var listLevelNFont = listLevelN.Font;
                listLevelNFont.Bold = True;
                listLevelNFont.Italic = False;
                listLevelNFont.Color = (Word.WdColor.wdColorAutomatic);
                listLevelNFont.Size = 14;
                listLevelNFont.Name = "Times New Roman";
                listLevelN.LinkedStyle = fonts[n - 1];
            }
            style.LinkToListTemplate(WordApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1], 1);
            return style;
        }
        private Word.Style GetDefaultStyle(string name = "Default")
        {
            var style = WordDoc.Styles.Add(name, Word.WdStyleType.wdStyleTypeParagraph);
            var font = style.Font;
            var paragraphFormat = style.ParagraphFormat;
            var shading = paragraphFormat.Shading;
            var borders = paragraphFormat.Borders;
            style.AutomaticallyUpdate = false;
            font.Name = "Times New Roman";
            font.Size = 14;
            font.Bold = False;
            font.Italic = False;
            font.Underline = Word.WdUnderline.wdUnderlineNone;
            font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            font.StrikeThrough = False;
            font.DoubleStrikeThrough = False;
            font.Outline = False;
            font.Emboss = False;
            font.Shadow = False;
            font.Hidden = False;
            font.SmallCaps = False;
            font.AllCaps = False;
            font.Color = Word.WdColor.wdColorAutomatic;
            font.Engrave = False;
            font.Superscript = False;
            font.Subscript = False;
            font.Scaling = 100;
            font.Kerning = 1;
            font.Animation = Word.WdAnimation.wdAnimationNone;
            font.Ligatures = Word.WdLigatures.wdLigaturesStandardContextual;
            font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            font.ContextualAlternates = 0;
            paragraphFormat.LeftIndent = WordApp.CentimetersToPoints(0);
            paragraphFormat.RightIndent = WordApp.CentimetersToPoints(0);
            paragraphFormat.SpaceBefore = 0;
            paragraphFormat.SpaceBeforeAuto = False;
            paragraphFormat.SpaceAfter = 0;
            paragraphFormat.SpaceAfterAuto = False;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            paragraphFormat.WidowControl = True;
            paragraphFormat.KeepWithNext = False;
            paragraphFormat.KeepTogether = False;
            paragraphFormat.PageBreakBefore = False;
            paragraphFormat.NoLineNumber = False;
            paragraphFormat.Hyphenation = True;
            paragraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(0);
            paragraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            paragraphFormat.CharacterUnitLeftIndent = 0;
            paragraphFormat.CharacterUnitRightIndent = 0;
            paragraphFormat.CharacterUnitFirstLineIndent = 0;
            paragraphFormat.LineUnitBefore = 0;
            paragraphFormat.LineUnitAfter = 0;
            paragraphFormat.MirrorIndents = False;
            paragraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            paragraphFormat.CollapsedByDefault = False;
            style.NoSpaceBetweenParagraphsOfSameStyle = false;
            paragraphFormat.TabStops.ClearAll();
            shading.Texture = Word.WdTextureIndex.wdTextureNone;
            shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            borders.DistanceFromTop = 1;
            borders.DistanceFromLeft = 4;
            borders.DistanceFromBottom = 1;
            borders.DistanceFromRight = 4;
            borders.Shadow = false;
            style.Frame.Delete();
            return style;
        }

    }
}
