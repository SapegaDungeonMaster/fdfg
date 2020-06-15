using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp4
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (comboBox1.SelectedIndex == 0)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 1)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 2)
            {
                Form3 f3= new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 3)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 4)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 5)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 6)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 7)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
            if (comboBox1.SelectedIndex == 8)
            {
                Form3 f3 = new Form3();
                f3.Text = comboBox1.SelectedItem.ToString();
                f3.Show();
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;
        }
    }
}
