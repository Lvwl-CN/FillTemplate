using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FillTemplate.Pages
{
    /// <summary>
    /// SheetSelect.xaml 的交互逻辑
    /// </summary>
    public partial class SheetSelect : Window
    {
        public List<string> Sheets { set { this.cbx.ItemsSource = value; } }

        public string Sheet { get { return this.cbx.SelectedItem.ToString(); } }
        public SheetSelect()
        {
            InitializeComponent();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}
