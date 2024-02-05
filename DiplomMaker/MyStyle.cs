using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public class MyStyle
    {
        public Word.Style Style { get; set; }
        public bool IsTopHeading => Style.ParagraphFormat.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel1;
        public bool IsHeading => Style.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText;
        public MyStyle(Word.Style style)
        {
            Style = style;
        }
    }
}
