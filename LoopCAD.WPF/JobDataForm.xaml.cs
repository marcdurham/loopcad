using System.Windows;
using System.Windows.Controls;

namespace LoopCAD.WPF
{
    /// <summary>
    /// Interaction logic for JobDataForm.xaml
    /// </summary>
    public partial class JobDataForm : Window
    {
        public JobDataForm(JobData data)
        {
            InitializeComponent();
            DataContext = data;
        }

        public JobData JobData { get; set; } = new JobData() { JobNumber = "EMPTY" };

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            //jobNumberBox.Text = JobData.JobNumber;
            //jobNameBox.Text = JobData.JobName;
        }

        private void jobNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void jobNumberBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            Close();
        }
    }
}
