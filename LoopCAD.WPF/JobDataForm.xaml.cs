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
    /// Interaction logic for JobDataForm.xaml
    /// </summary>
    public partial class JobDataForm : Window
    {
        public JobDataForm()
        {
            InitializeComponent();
        }

        public JobData JobData { get; set; } = new JobData();

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            jobNumberBox.Text = JobData.JobNumber;
            jobNameBox.Text = JobData.JobName;
        }

        private void jobNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void jobNumberBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
