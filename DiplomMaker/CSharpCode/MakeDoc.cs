using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    public partial class MakeDoc
    {
        const int True = -1;
        const int False = 0;
        /// <summary>
        /// Использован ли этот экземпляр класса
        /// </summary>
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
        /// <summary>
        /// Путь для сохранения файла по умолчанию
        /// </summary>
        public string Path { get; set; }
        public MakeDoc(string path = null)
        {
            Path = path;
            // init word things
            WordApp = new Word.Application() { Visible = true };
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
                WordApp.Quit(SaveChanges: false);
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
            DocContent doc = ParseText(text);
            // add text
            var selection = WordApp.Selection;
            for (int i = 0; i < doc.Paragraphs.Length; i++)
            {
                Paragraph par = doc.Paragraphs[i];
                par.PreAction(selection);
                par.MainAction(selection);
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
            while (true)
            {
                try
                {
                    WordDoc.SaveAs2(FileName: path ?? Path ?? "temp.docx");
                    WordApp.Quit();
                    break;
                }
                catch (Exception e)
                {
                    if (MessageBox.Show(e.Message + "\nYes - Продолжить попытки сохранения\nNo - не сохранять файл", "Exception", MessageBoxButton.YesNo) == MessageBoxResult.No)
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

        public DocContent ParseText(string text)
        {
            var doc = new DocContent();

            // Найти все формулы в тексте
            var formulasMatch = Paragraph.FormulaParagraph.Matches(text);
            for (int i = 0; i < formulasMatch.Count; i++)
            {
                // Найти номер в строке формулы
                var numberMatch = Paragraph.Numbers.Match(formulasMatch[i].Value);
                // Добавить номер в словарь номеров
                doc.Numbers[numberMatch.Value] = (i + 1).ToString();
            }
            var imagesMatch = Paragraph.ImageParagraph.Matches(text);
            for (int i = 0; i < imagesMatch.Count; i++)
            {
                // Найти номер в строке изображения
                var numberMatch = Paragraph.Numbers.Match(imagesMatch[i].Value);
                // Добавить номер в словарь номеров
                doc.Numbers[numberMatch.Value] = (i + 1).ToString();
            }
            // Найти все номера в тексте (кроме формул, картинок и таблиц)
            var numbersMatch = Paragraph.NumbersInAllTextExceptFormulasImages.Matches(text);
            foreach (Match numberMatch in numbersMatch)
            {
                // Преобразовать номер (убрать символ номера и лишние пустые символы)
                var key = numberMatch.Value.Replace("№", "").Trim();
                // Если этот номер есть в словаре, то
                if (doc.Numbers.ContainsKey(key))
                {
                    // Сохранить начальную позицию первого появления номера в тексте
                    var pos = text.IndexOf(numberMatch.Value);
                    // Заменить переменный номер на номер из словаря
                    text = text.Substring(0, pos) + doc.Numbers[key] + text.Substring(pos+numberMatch.Value.Length);
                    // В итоге будут заменены все номера, кроме тех, которые в строквх формул, изображений и таблиц,
                    // чтобы не сбивать форматирование последних.
                }
                
            }
            // Разделить весь текст по строкам
            var lines = text.Trim().Replace("\r\n", "\n").Split('\r', '\n');
            doc.Paragraphs = new Paragraph[lines.Length];
            // Для каждой строки
            for (int i = 0; i < lines.Length; i++)
            {
                doc.Paragraphs[i] = new Paragraph();
                var line = lines[i];
                // Если строка - обычный заголовок, то
                var match = Paragraph.RegularHeading.Match(line);
                if (match.Success)
                {
                    // Установить стиль "MainHeading"
                    doc.Paragraphs[i].MyStyle = new MyStyle(MainHeading);
                    doc.Paragraphs[i].Text = match.Value;
                    doc.Paragraphs[i].PreAction = (s) => s.InsertBreak(Word.WdBreakType.wdPageBreak);
                    continue;
                }
                // Если строка - списочный заголовок, то
                match = Paragraph.ListHeading.Match(line);
                if (match.Success)
                {
                    Word.Style style;
                    // Установить стиль в зависимости от количества решёток в начале строки
                    // Правда будет проблема, если # стоит не в начале. На позицию # пока нет проверки
                    int level = line.Count(x => x == '#') - 1;
                    if (i > 0 && doc.Paragraphs[i - 1].MyStyle.IsHeading)
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
                    doc.Paragraphs[i].MyStyle = new MyStyle(style);
                    doc.Paragraphs[i].Text = match.Value;
                    // Если заголовок первого уровня - вставить разрыв страницы
                    if (level == 0) doc.Paragraphs[i].PreAction = (s) => s.InsertBreak(Word.WdBreakType.wdPageBreak);
                    continue;
                }
                // Если строка является строкой списка чёрточек, то
                match = Paragraph.DashList.Match(line);
                if (match.Success)
                {
                    doc.Paragraphs[i].MyStyle = new MyStyle(DashList);
                    doc.Paragraphs[i].Text = match.Value;
                    continue;
                }
                // Если строка является строкой формулы, то
                match = Paragraph.FormulaParagraph.Match(line);
                if (match.Success)
                {
                    doc.Paragraphs[i].MyStyle = new MyStyle(FormulaParagraph);
                    // Найти саму формулу в этой строке
                    var formulaMatch = Paragraph.FormulaLine.Match(line);
                    // Найти номер этой формулы
                    var number = Paragraph.Numbers.Match(line);
                    // Изменить MainAction
                    doc.Paragraphs[i].MainAction = (s) =>
                    {
                        // Вставить начальную табуляцию
                        s.TypeText("\t");
                        // Добавить объект формулы
                        var oMath = s.OMaths.Add(s.Range);
                        // Ввести формулу
                        s.TypeText(formulaMatch.Value);
                        // Построить формулу
                        oMath.OMaths.BuildUp();
                        //// Передвинуть курсор на 1 позицию вправо (чтобы выйти из формулы)
                        //s.MoveRight(Word.WdUnits.wdCharacter, 1);
                        // Передвинуть курсор в конец строки
                        // Несколько раз, так как возникает какая-то брехня с границами формул,
                        // и 1 раза не достаточно.
                        s.EndKey(Word.WdUnits.wdLine);
                        s.EndKey(Word.WdUnits.wdLine);
                        s.EndKey(Word.WdUnits.wdLine);
                        // Ввести ещё 1 табуляцию и номер формулы
                        s.TypeText($"\t({doc.Numbers[number.Value]})");
                        // Установить стиль формулы
                        s.set_Style(FormulaParagraph);
                    };
                    continue;
                }
                match = Paragraph.ImageParagraph.Match(line);
                if (match.Success)
                {
                    var path = Paragraph.ImagePath.Match(match.Value);
                    var title = Paragraph.ImageTitle.Match(match.Value);
                    var number = Paragraph.Numbers.Match(match.Value);
                    doc.Paragraphs[i].MyStyle = new MyStyle(ImageParagraph);
                    doc.Paragraphs[i].MainAction = (s) =>
                    {
                        s.InlineShapes.AddPicture(path.Value, LinkToFile: false, SaveWithDocument: true);
                        s.set_Style(ImageParagraph);
                        s.MoveRight(Word.WdUnits.wdCharacter, 1);
                        s.TypeParagraph();
                        s.TypeText($"Рисунок {doc.Numbers[number.Value]} – {title.Value}");
                        s.set_Style(ImageNumberParagraph);
                    };
                    continue;
                }
                // Если ни один стиль не подошёл, то это обычная строка
                doc.Paragraphs[i].MyStyle = new MyStyle(RegularText);
                doc.Paragraphs[i].Text = line;
            }
            return doc;
        }

    }
}
