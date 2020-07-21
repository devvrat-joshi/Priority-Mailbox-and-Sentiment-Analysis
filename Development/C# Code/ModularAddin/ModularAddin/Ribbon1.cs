using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
namespace ModularAddin
{
    public partial class Ribbon1
    {
        public void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public void button2_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        public void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

       

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            MessageBox.Show("hello");
            Process.Start(path + "open_settings.py");
        }

        

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("OK");
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
