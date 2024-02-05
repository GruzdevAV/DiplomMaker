using Microsoft.VisualBasic;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.TextFormatting;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public partial class MakeDoc
    {
        const int True = -1;
        const int False = 0;
        private bool Used = false;
        Word.Application WordApp { get; set; }
        Word.Document WordDoc { get; set; }
        Word.Style MainHeading { get; set; }
        Word.Style ListHeading1 { get; set; }
        Word.Style ListHeading2AfterHead { get; set; }
        Word.Style ListHeading2AfterText { get; set; }
        Word.Style ListHeading3AfterHead { get; set; }
        Word.Style ListHeading3AfterText { get; set; }
        Word.Style ListHeading4AfterHead { get; set; }
        Word.Style ListHeading4AfterText { get; set; }
        Word.Style RegularText { get; set; }
        Word.Style FormulaParagraph { get; set; }
        Word.Style ImageParagraph { get; set; }
        Word.Style ImageNumberParagraph { get; set; }
        Word.Style TableNumberParagraph { get; set; }
        Word.Style AfterTableBeforeText { get; set; }
        Word.Style AfterTableBeforeHeading { get; set; }
        Word.Style TestListStyle { get; set; }
        Word.Style DashList { get; set; }
        public string Path { get; set; }
        public MakeDoc(string path = null)
        {
            Path = path;
            // init word things
            WordApp = new Word.Application();
            WordDoc = WordApp.Documents.Add();

            // Initialize Styles
            InitStyles();
        }
        ~MakeDoc()
        {
            if (Used) return;
            WordDoc.SaveAs2(FileName: Path ?? "temp.docx");
            WordApp.Quit();
        }
        public void AddText(string text)
        {
            if (Used)
            {
                RaiseUsedException();
                return;
            }
            // Parse text
            Paragraph[] pars = ParseText(text);
            // add text
            var selection = WordApp.Selection;
            for (int i = 0; i < pars.Length; i++)
            {
                Paragraph par = pars[i];
                par.PreAction(selection);
                selection.TypeText(par.Text);
                selection.set_Style(par.MyStyle.Style);
                par.PostAction(selection);
            }
        }

        private static void RaiseUsedException()
        {
            throw new Exception("Этот объект уже был использован (сохранён и закрыт)!\nСоздайте новый.");
        }

        public void SaveAndFinish(string path = null)
        {
            if (Used)
            {
                RaiseUsedException();
                return;
            }
            WordDoc.SaveAs2(FileName: path ?? Path ?? "temp.docx");
            WordApp.Quit();
            Used = true;
        }
        private void InitStyles()
        {
            MainHeading = GetStyleMainHeading();
            ListHeading1 = GetStyleListHeading("ListHeading1", 1, false);
            ListHeading2AfterHead = GetStyleListHeading("ListHeading2AfterHead", 2, false);
            ListHeading2AfterText = GetStyleListHeading("ListHeading2AfterText", 2, true);
            ListHeading2AfterText.set_BaseStyle(ListHeading2AfterHead.NameLocal);
            ListHeading3AfterText = GetStyleListHeading("ListHeading3AfterText", 3, true);
            ListHeading3AfterHead = GetStyleListHeading("ListHeading3AfterHead", 3, false);
            ListHeading3AfterText.set_BaseStyle(ListHeading3AfterHead.NameLocal);
            ListHeading4AfterHead = GetStyleListHeading("ListHeading4AfterHead", 4, false);
            ListHeading4AfterText = GetStyleListHeading("ListHeading4AfterText", 4, true);
            ListHeading4AfterText.set_BaseStyle(ListHeading4AfterHead.NameLocal);
            RegularText = GetStyleRegularText();
            FormulaParagraph = GetStyleFormulaParagraph();
            ImageParagraph = GetStyleImageParagraph();
            ImageNumberParagraph = GetStyleImageNumberParagraph();
            TableNumberParagraph = GetStyleTableNumberParagraph();
            AfterTableBeforeText = GetStyleAfterTable("AfterTableBeforeText", true);
            AfterTableBeforeHeading = GetStyleAfterTable("AfterTableBeforeHeading", false);
            TestListStyle = GetStyleForListHeadings();
            DashList = GetStyleDashList();
            ListHeading1.set_BaseStyle(TestListStyle.NameLocal);
            // Почему-то после style.set_BaseStyle(baseStyle) у style исчезает полужирность,
            // даже еслу у baseStyle она есть
            // Поэтому этот цикл нужен
            foreach (var style in new[] { ListHeading1, ListHeading2AfterText, ListHeading3AfterText, ListHeading4AfterText })
                style.Font.Bold = True;
        }

        public Paragraph[] ParseText(string text)
        {
            var lines = text.Trim().Replace("\r\n", "\n").Split('\r', '\n');
            var arr = new Paragraph[lines.Length];
            for (int i = 0; i < lines.Length; i++)
            {
                arr[i] = new Paragraph();
                var line = lines[i];
                var match = Paragraph.RegularHeading.Match(line);
                if (match.Success)
                {
                    arr[i].MyStyle = new MyStyle(MainHeading);
                    arr[i].Text = match.Value;
                    arr[i].PreAction = (s) => s.InsertBreak(Word.WdBreakType.wdPageBreak);
                    continue;
                }
                match = Paragraph.ListHeading.Match(line);
                if (match.Success)
                {
                    Word.Style style;
                    int level = line.Count(x => x == '#') - 1;
                    if (i > 0 && arr[i - 1].MyStyle.IsHeading)
                    {
                        style = (new[] {
                            ListHeading1,
                            ListHeading2AfterHead,
                            ListHeading3AfterHead,
                            ListHeading4AfterHead
                        }[level]);
                    }
                    else
                    {
                        style = (new[] {
                            ListHeading1,
                            ListHeading2AfterText,
                            ListHeading3AfterText,
                            ListHeading4AfterText
                        }[level]);
                    }
                    arr[i].MyStyle = new MyStyle(style);
                    arr[i].Text = match.Value;
                    if (level == 0) arr[i].PreAction = (s) => s.InsertBreak(Word.WdBreakType.wdPageBreak);
                    continue;
                }
                match = Paragraph.DashList.Match(line);
                if (match.Success)
                {
                    arr[i].MyStyle = new MyStyle(DashList);
                    arr[i].Text = match.Value;
                    continue;
                }
                arr[i].MyStyle = new MyStyle(RegularText);
                arr[i].Text = line;
            }
            return arr;
        }

    }
}
