using Word = Microsoft.Office.Interop.Word;
using System;
using System.Text.RegularExpressions;

namespace DiplomMaker
{
    public class Paragraph
    {
        public static readonly Regex RegularHeading = new Regex(@"(?<=^\s*?@\s*?)[^@]+?(?=\s*?)$", RegexOptions.Multiline);
        //public static readonly Regex RegularHeading = new Regex(@"(?<=^\s*?@{1,4}\s*?)[^@]+?(?=\s*?)$", RegexOptions.Multiline);
        public static readonly Regex ListHeading = new Regex(@"(?<=^\s*?#{1,4}\s*?)[^#]+?(?=\s*?)$$", RegexOptions.Multiline);
        // -–—− ←               это разные симолы                       ↓
        public static readonly Regex DashList = new Regex(@"(?<=^\s*?[-–—−]\s*?).+?(?=\s*?)$", RegexOptions.Multiline);
        private string _text;

        public Action<Word.Selection> PreAction { get; set; } = (_) => { };
        public Action<Word.Selection> PostAction { get; set; } = (s) => { s.TypeParagraph(); };
        public MyStyle MyStyle { get; set; }
        public string Text { get => _text; set => _text = value.Trim(); }
    }
}
