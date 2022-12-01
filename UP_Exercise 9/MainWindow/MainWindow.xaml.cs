using NetOffice.WordApi.Enums;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using UP_Exercise_9.ApplicationLogic;

namespace UP_Exercise_9
{
    public partial class MainWindow : Window
    {
        private static int Cols = 0;
        private static int Rows = 0;
        public MainWindow()
        {
            InitializeComponent();
            Closing += ShowCloseMessage;
            rows.Num.TextChanged += RowsTextChanged;
            cols.Num.TextChanged += ColsTextChanged;
        }
        private void ShowCloseMessage(object sender, CancelEventArgs e)
        {
            if (ActionConfirmation("Вы уверены, что хотите закрыть приложение?") == MessageBoxResult.No) e.Cancel = true;
            else WordClass.Application.Quit(WdSaveOptions.wdDoNotSaveChanges);
        }
        private static MessageBoxResult ActionConfirmation(string question) => MessageBox.Show(question, "Подтвердите действие", MessageBoxButton.YesNo, MessageBoxImage.Question);
        private void RowsTextChanged(object sender, TextChangedEventArgs e)
        {
            Rows = Convert.ToInt32(CheckTextBoxValue(rows.Num.Text, 0, 50, (TextBox)sender));
            GenerateTable();
        }
        private void ColsTextChanged(object sender, TextChangedEventArgs e)
        {
            Cols = Convert.ToInt32(CheckTextBoxValue(cols.Num.Text, 0, 10, (TextBox)sender));
            GenerateTable();
        }
        private static string CheckTextBoxValue(string value, int min, int max, TextBox tb)
        {
            if (value == "") return min.ToString();
            if (Convert.ToInt32(value) < min) tb.Text = min.ToString();
            else if (Convert.ToInt32(value) > max) tb.Text = max.ToString();
            else return value;
            return tb.Text;
        }
        private void GenerateTable()
        {
            TableGrid.Children.Clear();
            if (Cols <= 0 || Rows <= 0) return;
            StackPanel RowPanel = new() { Orientation = Orientation.Vertical };
            for (int i = 0; i < Rows; i++)
            {
                StackPanel Row = new() { Orientation = Orientation.Horizontal };
                for (int j = 0; j < Cols; j++)
                {
                    TextBox tb = new()
                    {
                        Height = TableGrid.Height / Rows,
                        Width = TableGrid.Width / Cols,
                        Style = (Style)Application.Current.Resources["TextBoxStyle"],
                        FontSize = Math.Max(10, 80 - (Cols + Rows) * 5)
                    };
                    Row.Children.Add(tb);
                }
                RowPanel.Children.Add(Row);
            }
            TableGrid.Children.Add(RowPanel);
        }
        private void NamePreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if ("{}&%~#+=[]:*?;«\"/\\<>|".Contains(e.Text)) e.Handled = true;
        }
        private void CreateTableClick(object sender, RoutedEventArgs e)
        {
            if (ActionConfirmation($"Вы уверены, что хотите создать файл {Name.Text}.docx с этой таблицей?") == MessageBoxResult.No) return;
            Stopwatch stopwatch = new();
            stopwatch.Start();
            Exception? ex = WordClass.CreateWordTable(Name.Text, TableGrid, Rows, Cols);
            stopwatch.Stop();
            if (ex == null)
            {
                MessageBox.Show($"Файл с таблицей успешно создан!\nВремя выполнения: {stopwatch.Elapsed.TotalSeconds:f2} секунд", "Создание файла", MessageBoxButton.OK, MessageBoxImage.Information); ;
            }
            else
            {
                MessageBox.Show($"Ошибка при создании файла:\n{ex.Message}\nВремя выполнения: {stopwatch.Elapsed.TotalSeconds:f2} секунд", "Создание файла", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
