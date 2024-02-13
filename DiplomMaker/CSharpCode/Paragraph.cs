using Word = Microsoft.Office.Interop.Word;
using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public class Paragraph
    {
        public static readonly Regex RegularHeading = new Regex(@"(?<=^\s*?@\s*?)[^@]+?(?=\s*?)$", RegexOptions.Multiline);
        //public static readonly Regex RegularHeading = new Regex(@"(?<=^\s*?@{1,4}\s*?)[^@]+?(?=\s*?)$", RegexOptions.Multiline);
        public static readonly Regex ListHeading = new Regex(@"(?<=^\s*?#{1,4}\s*?)[^#]+?(?=\s*?)$", RegexOptions.Multiline);
        // -–—− ←               это разные симолы                       ↓
        public static readonly Regex DashList = new Regex(@"(?<=^\s*?[-–—−]\s*?).+?(?=\s*?)$", RegexOptions.Multiline);
        // Выделяет всю строку с формулой
        public static readonly Regex FormulaParagraph = new Regex(@"^\s*?!\s*?formula\s*?\(.+?\)\s*?;\s*?№\s*?[\wа-яё\d]+?\s*?$",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        public static readonly Regex FormulaLine = new Regex(@"(?<=^\s*?!\s*?formula\s*?\().+?(?=\)\s*?;\s*?№\s*?[\wа-яё\d]+?\s*?$)",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        // Выделяет номер из строк формулы и картинок. Лишь "N" из "№ N" или "№N"
        public static readonly Regex Numbers = new Regex(@"(?<=^\s*?!\s*?((formula)|(img))\s*?.+?№\s*?)[\wа-яё\d]+?(?=\s*?$)",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        // Выделяет номер во всём тексте
        public static readonly Regex NumbersInAllText = new Regex(@"№\s*?[\wа-яё\d]+",
            RegexOptions.Multiline | RegexOptions.IgnoreCase);
        // Выделяет номер во всём тексте, кроме формул и изображений
        public static readonly Regex NumbersInAllTextExceptFormulasImages = new Regex(@"(?<!^\s*?!\s*?((formula)|(img))\s*?.+?)№\s*?[\wа-яё\d]+",
            RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private string _text;

        public Paragraph() => MainAction = (s) =>
        {
            s.TypeText(Text);
            s.set_Style(MyStyle.Style);
        };

        public Action<Word.Selection> PreAction { get; set; } = (_) => { };
        public Action<Word.Selection> MainAction { get; set; }
        public Action<Word.Selection> PostAction { get; set; } = (s) => s.TypeParagraph();
        public MyStyle MyStyle { get; set; }
        public string Text { get => _text; set => _text = value.Trim(); }
    }
}
