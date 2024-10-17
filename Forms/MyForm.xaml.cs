using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
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
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace RAA_Level2
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    /// Step 3  : Code behind
    public partial class MyForm : Window
    {
        public MyForm()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            tbx.Text = "Select an excel file";
            // open dialog to select file
            OpenFileDialog openfile = new OpenFileDialog();
            // give the opening space an inital directory
            openfile.InitialDirectory = "C:\\";
            //filter what is allowed to be select aka .csv
            openfile.Filter = "excel files(*.xlsx)|*.xlsx";
            //if statement to say, if file is select then copy name to text

            if (openfile.ShowDialog() == true)
            {
                //copy text to tbx.text
                tbx.Text = openfile.FileName;
            }
            else
            {
                tbx.Text = "";
            }
        }
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        public string getTextboxValue()
        {
            return tbx.Text;
        }        

        public bool getChbFloorPlans()
        {
            if(chbFloorPlans.IsChecked==true)
                return true;
            else return false;
        }

        public bool getChbCeilingPlans()
        {
            if(chbCeilingPlans.IsChecked==true) 
                return true;
            else return false;
        }

        public string getUnitGroup()
        {
            if (rbImperial.IsChecked == true) return rbImperial.ToString();
            else return rbMetric.ToString();
        }

    }
}
