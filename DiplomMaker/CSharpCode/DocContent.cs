using System.Collections.Generic;

namespace DiplomMaker
{
    /// <summary>
    /// Содержимое документа
    /// </summary>
    public class DocContent
    {
        /// <summary>
        /// Массив параграфов (текст со стилями и всё такое)
        /// </summary>
        public Paragraph[] Paragraphs { get; set; }
        /// <summary>
        /// Словарь номеров (хранит номера всех рисунков, таблиц и формул, 
        /// чтобы можно было заменить "см. рис. ыы, формулу smthing" на "см. рис. 1, формулу 2" и т.п.
        /// </summary>
        public Dictionary<string, string> Numbers { get; } = new Dictionary<string, string>(new IgnoreCaseComparer());
    }
}
