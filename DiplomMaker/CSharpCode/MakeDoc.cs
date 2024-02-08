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
using System.Windows;
using System.Windows.Controls;
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
            // Путь для сохранения файла по умолчанию
            Path = path;

            // init word things
            WordApp = new Word.Application();
            WordDoc = WordApp.Documents.Add();

            // Initialize
            InitStyles();
            // Настройка отствупов полей
            SetPageMargin();
            // Настройка нумерации страниц
            SetPageNumbers();
        }

        private void SetPageNumbers()
        {
            // Вставлю разрыв раздела со следующей страницы,
            // а также заготавливаю места для титульника и содержания
            var selection = WordApp.Selection;
            selection.TypeText("Место для титульника");
            selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            selection.TypeText("Место для оглавления. Там дальше должен появиться перенос строки, так как следующий параграф должен быть заголовком первого уровня \"ВВЕДЕНИЕ\", а перед заголовками первого уровня ставится перенос страницы.");

            // Настраиваю номер страницы для втоорого раздела
            var footer = WordDoc.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            footer.LinkToPrevious = false;
            var pageNumbers = footer.PageNumbers;
            pageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabic;
            pageNumbers.StartingNumber = 2;

            // Добавить сам номер страницы
            var pageNumber = footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
            
            // Перехожу к текущему нижнему колонтитулу
            WordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            // Передвигаю курсор на 2 элемента влево (эквивалентно нажатию 2 раза на кнопку ←)
            // Это нужно, чтобы выделить тот странный объект, который создался при создании pageNumber
            selection.MoveLeft(Word.WdUnits.wdCharacter, 2);
            // Настраиваю стиль номера страницы
            // Можно, конечно, и не копировать стиль номера страницы, а установить стиль RegularText,
            // потом установить выравнивание по центру и удалить появившейся откуда-то перенос строки ниже номера,
            // но я решил оставить тот странный плавающий объект номера страницы
            var style = (Word.Style)selection.get_Style();
            var font = style.Font;
            //var paragraphFormat = style.ParagraphFormat;
            font.Name = "Times New Roman";
            font.Size = 14;
            font.Color = Word.WdColor.wdColorAutomatic;
            //paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;

            // Устанавливаю стиль для номера страницы
            selection.set_Style(style);

            // Возвращаюсь к главному документу
            WordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        private void SetPageMargin()
        {
            // Настройка номера страницы "отсюда" (где сейчас курсор стоит) до конца документа
            var pageSetup = WordDoc.Range(WordApp.Selection.Start, WordDoc.Content.End).PageSetup;
            // Вертикальная ориентация страницы
            pageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Настройка длины полей сверху, снизу, слева и справа соответственно.
            pageSetup.TopMargin = WordApp.CentimetersToPoints(2);
            pageSetup.BottomMargin = WordApp.CentimetersToPoints(2);
            pageSetup.LeftMargin = WordApp.CentimetersToPoints(3);
            pageSetup.RightMargin = WordApp.CentimetersToPoints(1.5f);
        }

        ~MakeDoc()
        {
            if (Used) return;
            try
            {
                WordDoc.SaveAs2(FileName: Path ?? "temp.docx");
            }
            finally
            {
                WordApp.Quit(SaveChanges:false);
            }
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
            while(true)
            {
                try
                {
                    WordDoc.SaveAs2(FileName: path ?? Path ?? "temp.docx");
                    WordApp.Quit();
                    break;
                }
                catch(Exception e) 
                {
                    if (MessageBox.Show(e.Message+"\nYes - Продолжить попытки сохранения\nNo - не сохранять файл", "Exception", MessageBoxButton.YesNo) == MessageBoxResult.No)
                    {
                        WordApp.Quit(false);
                        break;
                    }
                }
            }
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
