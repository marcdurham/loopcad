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

namespace LoopCAD.WPF
{
    /// <summary>
    /// Interaction logic for ChooseFloor.xaml
    /// </summary>
    public partial class ChooseFloor : Window
    {
        public string ChosenName { get; set; }
        public List<FloorTag> FloorTags { get; set; }

        public ChooseFloor()
        {
            InitializeComponent();
        }

        public ChooseFloor(List<FloorTag> floorTags)
        {
            FloorTags = floorTags;
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                var item = e.AddedItems[0] as ListBoxItem;
                ChosenName = item.Content as string;
            }

            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            foreach (var floorTag in FloorTags)
            {
                floorsList.Items.Add(floorTag.Name);
            }
        }
    }
}
