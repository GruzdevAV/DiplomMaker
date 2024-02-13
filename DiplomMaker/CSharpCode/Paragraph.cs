using Word = Microsoft.Office.Interop.Word;
using System;
using System.Text.RegularExpressions;

namespace DiplomMaker
{
    public class Paragraph
    {
        public static Regex RegularHeading { get; } = new Regex(@"(?<=^\s*?@\s*?)[^@]+?(?=\s*?)$", RegexOptions.Multiline);
        public static  Regex ListHeading { get; } = new Regex(@"(?<=^\s*?#{1,4}\s*?)[^#]+?(?=\s*?)$", RegexOptions.Multiline);
        // -–—− ←               это разные символы                       ↓
        public static  Regex DashList { get; } = new Regex(@"(?<=^\s*?[-–—−]\s*?).+?(?=\s*?)$", RegexOptions.Multiline);
        // Выделяет всю строку с формулой
        public static Regex FormulaParagraph { get; } = new Regex(@"^\s*?!\s*?formula\s*?\(.+?\)\s*?;\s*?№\s*?[\wа-яё\d]+?\s*?$",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        public static  Regex FormulaLine { get; } = new Regex(@"(?<=^\s*?!\s*?formula\s*?\().+?(?=\)\s*?;\s*?№\s*?[\wа-яё\d]+?\s*?$)",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        // Выделяет всю строку с картинкой
        public static Regex ImageParagraph { get; } = new Regex(@"!\s*?img\("".+?"";.+?\);\s*?№\s*?[\wа-яё\d]+?\s*?$",
        RegexOptions.Multiline | RegexOptions.IgnoreCase);
        public static Regex ImagePath { get; } = new Regex(@"(?<=!\s*?img\("").+?(?="".+?$)",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
        public static Regex ImageTitle { get; } = new Regex(@"(?<=!\s*?img.+?"";).+?(?=\s*\);.+?$)",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);

        // Выделяет номер из строк формулы и картинок. Лишь "N" из "№ N" или "№N"
        public static  Regex Numbers { get; } = new Regex(@"(?<=^\s*?!\s*?((formula)|(img))\s*?.+?№\s*?)[\wа-яё\d]+?(?=\s*?$)",
            RegexOptions.Multiline|RegexOptions.IgnoreCase);
        // Выделяет номер во всём тексте
        public static  Regex NumbersInAllText { get; } = new Regex(@"№\s*?[\wа-яё\d]+",
            RegexOptions.Multiline | RegexOptions.IgnoreCase);
        // Выделяет номер во всём тексте, кроме формул и изображений
        public static  Regex NumbersInAllTextExceptFormulasImages { get; } = new Regex(@"(?<!^\s*?!\s*?((formula)|(img))\s*?.+?)№\s*?[\wа-яё\d]+",
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
