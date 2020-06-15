using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Professional\Desktop\курсач.accdb");
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            login();
        }
        private void login()
        {
            string log = textBox1.Text;
            string pas = textBox2.Text;
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT count(*) FROM Преподаватель where Логин = '" + log + "' and Пароль ='" + pas + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            if (int.Parse(dt.Rows[0][0].ToString()) == 1)
            {
                Form2 f2 = new Form2();
                f2.Show();
                con.Close();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Не удается войти", "Ошибка",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox1.Checked == true)
            {
                textBox2.PasswordChar = '\0';
            }
            if (checkBox1.Checked == false)
            {
                textBox2.PasswordChar = '*';
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                login();
            }
        }
    }
}
