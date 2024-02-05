using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace DiplomMaker
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        FileDialog _markup_file;
        SaveFileDialog _word_file;
        Timer _update_timer;
        DateTime _update_time;
        string MarkupText
        {
            get
            {
                return new TextRange
                    (
                    rtb_markup_text.Document.ContentStart,
                    rtb_markup_text.Document.ContentEnd
                    ).Text;
            }
            set
            {
                rtb_markup_text.Document.Blocks.Clear();
                foreach (var paragraph in value.Replace("\r\n", "\n").Split('\n'))
                {
                    rtb_markup_text.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(paragraph)));
                }
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            _update_timer = new System.Timers.Timer
            {
                AutoReset = false,
                Interval = 2000,
                Enabled = false
            };
            _update_timer.Elapsed += _update_timer_Elapsed;
        }
        private void LoadMarkupText(string path)
        {
            try
            {
                MarkupText = File.ReadAllText(path, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }
        private void SaveMarkupText(string path)
        {
            try
            {
                using (var file = File.CreateText(path))
                {
                    file.WriteLineAsync(MarkupText);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }

        private void rtb_markup_text_TextChanged(object sender, TextChangedEventArgs e)
        {
            //_update_timer?.Stop();
            //_update_timer?.Start();
        }
        /// <summary>
        /// Нужен для обновления стиля текста разметки
        /// Но пока не работает из-за того, что я сделал это не в основном потоке
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _update_timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            //MarkupText = DateTime.Now.ToString();
        }

        private void btn_markup_choose_Click(object sender, RoutedEventArgs e)
        {
            _markup_file = new OpenFileDialog { Filter = "*.txt|*.txt|Все файлы (*.*)|*.*" };
            if (_markup_file.ShowDialog() != true)
                return;
            tbox_markup_path.Text = _markup_file.FileName;
            LoadMarkupText(tbox_markup_path.Text);
        }


        private void btn_markup_load_Click(object sender, RoutedEventArgs e)
        {
            if (tbox_markup_path.Text == string.Empty)
            {
                btn_markup_choose_Click(sender, e);
                return;
            }
            LoadMarkupText(tbox_markup_path.Text);
        }

        private void btn_markup_save_Click(object sender, RoutedEventArgs e)
        {
            if (tbox_markup_path.Text == string.Empty)
            {
                btn_markup_save_as_Click(sender, e);
                return;
            }
            SaveMarkupText(tbox_markup_path.Text);
        }

        private void btn_markup_save_as_Click(object sender, RoutedEventArgs e)
        {
            _markup_file = new SaveFileDialog { Filter = "*.txt|*.txt|Все файлы (*.*)|*.*" };
            if (_markup_file.ShowDialog() != true)
                return;
            tbox_markup_path.Text = _markup_file.FileName;
            SaveMarkupText(tbox_markup_path.Text);
        }


        private async void btn_word_save_Click(object sender, RoutedEventArgs e)
        {
            if (tbox_word_path.Text == string.Empty)
            {
                btn_word_save_as_Click(sender, e);
                return;
            }
            await SaveWordDoc(tbox_word_path.Text);

        }

        private async void btn_word_save_as_Click(object sender, RoutedEventArgs e)
        {
            _word_file = new SaveFileDialog { Filter = "*.doc|*.doc|*.docx|*.docx|Все файлы (*.*)|*.*" };
            if (_word_file.ShowDialog() != true)
                return;
            tbox_word_path.Text = _word_file.FileName;
            await SaveWordDoc(tbox_word_path.Text);
        }

        private async Task SaveWordDoc(string fileName)
        {
            MessageBox.Show("Начинаю сохранение Word-файла.", "Ожидайте");

            await Task.Run(() =>
            {
                MakeDoc doc = new MakeDoc(fileName);
                doc.AddText(MarkupText);
                doc.SaveAndFinish();
                MessageBox.Show("Файл Word сохранён.", "Успех");
            });
        }
        //Видеть документ https://stackoverflow.com/questions/1859641/load-word-excel-into-wpf
    }
}
