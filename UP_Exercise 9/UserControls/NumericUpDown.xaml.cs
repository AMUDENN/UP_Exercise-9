using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace UP_Exercise_9
{
    public partial class NumericUpDown : UserControl
    {
        public NumericUpDown()
        {
            InitializeComponent();
        }
        private void UpButtonClick(object sender, RoutedEventArgs e)
        {
            if (Convert.ToDouble(Num.Text) < int.MaxValue) Num.Text = (Convert.ToInt32(Num.Text) + 1).ToString();
        }
        private void DownButtonClick(object sender, RoutedEventArgs e)
        {
            if (Convert.ToDouble(Num.Text) > int.MinValue) Num.Text = (Convert.ToInt32(Num.Text) - 1).ToString();
        }
        private void NumPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!"0123456789".Contains(e.Text) || Num.Text.Length >= 2) e.Handled = true;
        }
    }
}
