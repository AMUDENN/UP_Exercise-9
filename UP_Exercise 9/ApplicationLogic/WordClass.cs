using System;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.IO;
using System.Windows.Controls;

namespace UP_Exercise_9.ApplicationLogic
{
    internal class WordClass
    {
        private static Word.Application _app = new();
        private static Word.Tools.Contribution.CommonUtils _utils = new(_app);
        public static Word.Application Application => _app;
        public static Exception? CreateWordTable(string title, Grid grid, int rows, int cols)
        {
            try
            {
                string name = _utils.File.Combine(Environment.CurrentDirectory, title, Word.Tools.Contribution.DocumentFormat.Normal);
                FileInfo file = new(name);
                if (file.Exists)
                    return new Exception("Файл с таким названием уже существует!");

                Word.Document doc = _app.Documents.Add();
                _app.Selection.TypeText("Таблица создана с помощью приложения на C#\n");
                Word.Table table = doc.Tables.Add(_app.Selection.Range, rows, cols);
                table.Borders.Enable = 1;

                StackPanel RowPanel = (StackPanel)grid.Children[0];
                for (int row = 0; row < rows; ++row)
                {
                    StackPanel Panel = (StackPanel)RowPanel.Children[row];
                    for (int column = 0; column < cols; ++column)
                    {
                        table.Cell(row + 1, column + 1).Select();
                        _app.Selection.TypeText(((TextBox)Panel.Children[column]).Text);
                    }
                }

                _app.Selection.HomeKey(WdUnits.wdStory, WdMovementType.wdExtend);
                _app.Selection.Font.Size = 14;
                _app.Selection.Font.Name = "Times New Roman";

                doc.SaveAs(name);
            }
            catch (Exception ex)
            {
                return ex;
            }
            return null;
        }
    }
}
