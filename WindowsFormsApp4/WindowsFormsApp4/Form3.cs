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
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace WindowsFormsApp4
{
    public partial class Form3 : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Professional\Desktop\курсач.accdb");
        public Form3()
        {
            InitializeComponent();

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = new OleDbCommand("Select дата_расписания FROM Расписание", con);
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                comboBox5.Items.Add(reader.GetString(0));
            }
        }

        private void группаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Группа";
            f3.Show();
            this.Close();
        }

        private void дисциплиныToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Дисциплины";
            f3.Show();
            this.Close();
        }

        private void дисциплиныгурппыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Дисциплины_группы";
            f3.Show();
            this.Close();
        }

        private void занятиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Занятие";
            f3.Show();
            this.Close();
        }

        private void кабинетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Кабинет";
            f3.Show();
            this.Close();
        }

        private void преподавательToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Преподаватель";
            f3.Show();
            this.Close();
        }

        private void расписаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Расписание";
            f3.Show();
            this.Close();
        }

        private void специальностьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Специальность";
            f3.Show();
            this.Close();
        }

        private void учащиесяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Учащиеся";
            f3.Show();
            this.Close();
        }
        public void loadTable(string selectTable)
        {
            if(con.State != ConnectionState.Open)
            {
                con.Open();
            }
            
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = selectTable;
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void Form3_Shown(object sender, EventArgs e)
        {

        }

        private void Form3_Shown_1(object sender, EventArgs e)
        {
            if (this.Text == "Группа")
            {
                loadTable(Queries.selectgrup);
                dataGridView1.Columns[0].Visible = false;

                string[] array = getSpecial().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
            }
            if (this.Text == "Дисциплины")
            {
                loadTable(Queries.selectdis);
                dataGridView1.Columns[0].Visible = false;


            }
            if (this.Text == "Дисциплины_группы")
            {
                loadTable(Queries.selectdis_grup);
                dataGridView1.Columns[0].Visible = false;

                string[] array = getDis().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
                string[] array1 = getGrup().Select(n => n.ToString()).ToArray();
                comboBox2.Items.AddRange(array1);
                string[] array2 = getKabinet().Select(n => n.ToString()).ToArray();
                comboBox3.Items.AddRange(array2);
            }
            if (this.Text == "Занятие")
            {
                loadTable(Queries.selectzaniat);
                string[] array = getKabinet().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
                string[] array1 = getRaspis().Select(n => n.ToString()).ToArray();
                comboBox2.Items.AddRange(array1);
                string[] array2 = getDis().Select(n => n.ToString()).ToArray();
                comboBox3.Items.AddRange(array2);
                string[] array3 = getPrepod().Select(n => n.ToString()).ToArray();
                comboBox4.Items.AddRange(array3);
                dataGridView1.Columns[0].Visible = false;

            }
            if (this.Text == "Кабинет")
            {
                loadTable(Queries.selectkabinet);
                dataGridView1.Columns[0].Visible = false;

            }
            if (this.Text == "Преподаватель")
            {
                loadTable(Queries.selectprepod);
                dataGridView1.Columns[0].Visible = false;

                string[] array = getDis().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
            }
            if (this.Text == "Расписание")
            {
                loadTable(Queries.selectraspisan);
                dataGridView1.Columns[0].Visible = false;

                string[] array = getKabinet().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
                string[] array1 = getGrup().Select(n => n.ToString()).ToArray();
                comboBox2.Items.AddRange(array1);
                string[] array2 = getDis().Select(n => n.ToString()).ToArray();
                comboBox3.Items.AddRange(array2);
            }
            if (this.Text == "Специальность")
            {
                loadTable(Queries.selectspecial);
                dataGridView1.Columns[0].Visible = false;

            }
            if (this.Text == "Учащиеся")
            {
                loadTable(Queries.selectycha);
                dataGridView1.Columns[0].Visible = false;

                string[] array = getGrup().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
            }
        }
        private int getIdByGrup(string nameGrup)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_группы FROM Группа where группа = '" + nameGrup + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdBySpecial(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_специальности FROM Специальность where название_специальности = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdByPrepod(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_преподавателя FROM Преподаватель where ФИО = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdByKabinet(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_кабинета FROM Кабинет where номер_кабинета = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdByDisGrup(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_дисциплины_группы FROM Дисциплины_группы where название_дисциплины = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdByDis(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT код_дисциплины FROM Дисциплины where название_дисциплины = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private int getIdByRaspis(string nameClient)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT id_расписания FROM Расписание where дата_расписания = '" + nameClient + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getGrup()
        {
            List<string> grup = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Группа";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                grup.Add(item[3].ToString());
            }
            return grup;
        }
        private List<string> getSpecial()
        {
            List<string> special = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Специальность";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                special.Add(item[1].ToString());
            }
            return special;
        }
        private List<string> getPrepod()
        {
            List<string> prepod = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Преподаватель";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                prepod.Add(item[2].ToString());
            }
            return prepod;
        }
        private List<string> getKabinet()
        {
            List<string> kabinet = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Кабинет";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                kabinet.Add(item[1].ToString());
            }
            return kabinet;
        }
        private List<string> getDisGrup()
        {
            List<string> disgrup = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Дисциплины_группы";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                disgrup.Add(item[3].ToString());
            }
            return disgrup;
        }
        private List<string> getDis()
        {
            List<string> dis = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Дисциплины";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                dis.Add(item[1].ToString());
            }
            return dis;
        }
        private List<string> getRaspis()
        {
            List<string> raspis = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Расписание";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                raspis.Add(item[1].ToString());
            }
            return raspis;
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.Text == "Группа")
            {
                int idSpecial = getIdBySpecial(comboBox1.SelectedItem.ToString());
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Группа (количество_учащихся, код_специальности, группа) VALUES('" + textBox2.Text + "'," + idSpecial + ",'" + textBox3.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_группы, количество_учащихся, название_специальности, группа FROM Группа, Специальность WHERE Группа.код_специальности=Специальность.код_специальности ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Дисциплины")
            {
                bool result = checkBox1.Checked;
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Дисциплины (название_дисциплины,экзамен) VALUES('" + textBox2.Text + "'," + result + ")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Дисциплины";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            else if (this.Text == "Дисциплины_группы")
            {
                int idDis = getIdByDis(comboBox1.SelectedItem.ToString());
                int idGrup = getIdByGrup(comboBox2.SelectedItem.ToString());
                int idKabinet = getIdByKabinet(comboBox3.SelectedItem.ToString());
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Дисциплины_группы (код_дисциплины,код_группы,код_кабинета,количество_часов_дисциплины) VALUES(" + idDis + "," + idGrup + "," + idKabinet + ",'" + textBox2.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_дисциплины_группы, название_дисциплины, группа, номер_кабинета,количество_часов_дисциплины FROM Дисциплины_группы, Дисциплины, Группа, Кабинет WHERE Дисциплины_группы.код_дисциплины=Дисциплины.код_дисциплины and Дисциплины_группы.код_группы=Группа.код_группы and Дисциплины_группы.код_кабинета=Кабинет.код_кабинета";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Занятие")
            {
                int idKabinet = getIdByKabinet(comboBox1.SelectedItem.ToString());
                int idRaspis = getIdByRaspis(comboBox2.SelectedItem.ToString());
                int idDis = getIdByDis(comboBox3.SelectedItem.ToString());
                int idPrepod = getIdByPrepod(comboBox4.SelectedItem.ToString());
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Занятие (код_кабинета,группа,id_расписания,код_дисциплины,код_преподавателя) VALUES(" + idKabinet + ",'" + textBox2.Text + "'," + idRaspis + "," + idDis + "," + idPrepod + ")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_занятия, номер_кабинета, группа, дата_расписания,ФИО,название_дисциплины FROM Занятие, Кабинет, Расписание,Дисциплины,Преподаватель WHERE Занятие.код_кабинета=Кабинет.код_кабинета and Занятие.id_расписания=Расписание.id_расписания and Занятие.код_дисциплины=Дисциплины.код_дисциплины and Занятие.код_преподавателя=Преподаватель.код_преподавателя";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Кабинет")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Кабинет (номер_кабинета) VALUES('" + textBox2.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Кабинет";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Преподаватель")
            {



                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Преподаватель (номер_кабинета_преподавателя,ФИО) VALUES('" + textBox2.Text + "','" + textBox3.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectprepod;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                loadTable(Queries.selectprepod);
            }
            else if (this.Text == "Расписание")
            {
                int idKabinet = getIdByKabinet(comboBox1.SelectedItem.ToString());
                int idGruppi = getIdByGrup(comboBox2.SelectedItem.ToString());
                int idDisc = getIdByDis(comboBox3.SelectedItem.ToString());
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Расписание (дата_расписания,код_кабинета, код_группы, код_дисципины) VALUES('" + textBox2.Text + "'," + idKabinet + ", "+idGruppi+", "+idDisc+")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectraspisan;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Специальность")
            {

                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Специальность (название_специальности) VALUES('" + textBox2.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Специальность";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (this.Text == "Учащиеся")
            {

                int idGrup = getIdByGrup(comboBox1.SelectedItem.ToString());
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Учащиеся (ФИО,телефон,улица,дата_рождения,код_группы) VALUES('" + textBox2.Text + "'," + textBox3.Text + ",'" + textBox4.Text + "','" + textBox5.Text + "'," + idGrup + ")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Учащиеся.код_учащегося, Учащиеся.ФИО, Учащиеся.телефон, Учащиеся.улица, Учащиеся.дата_рождения, Группа.группа FROM Группа INNER JOIN Учащиеся ON Группа.код_группы = Учащиеся.код_группы";

                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form3_Shown(sender, e);
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.Text == "Группа")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Группа WHERE код_группы=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Группа";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Дисциплины")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Дисциплины WHERE код_дисциплины=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Дисциплины";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Дисциплины_группы")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Дисциплины_группы WHERE код_дисциплины_группы=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_дисциплины_группы, название_дисциплины, группа, номер_кабинета,количество_часов_дисциплины FROM Дисциплины_группы, Дисциплины, Группа, Кабинет WHERE Дисциплины_группы.код_дисциплины=Дисциплины.код_дисциплины and Дисциплины_группы.код_группы=Группа.код_группы and Дисциплины_группы.код_кабинета=Кабинет.код_кабинета";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Занятие")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Занятие WHERE код_занятия=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Занятие";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Кабинет")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Кабинет WHERE код_кабинета=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Кабинет";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Преподаватель")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Преподаватель WHERE код_преподавателя=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Преподаватель";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Расписание")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Расписание WHERE id_расписания=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectraspisan;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Специальность")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Специальность WHERE код_специальности=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Специальность";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Учащиеся")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE From Учащиеся WHERE код_учащегося=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Учащиеся";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.Text == "Группа")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                int idSpecial = getIdBySpecial(comboBox1.Text);
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Группа SET количество_учащихся=" + textBox2.Text + ", код_специальности=" + idSpecial + " , группа='" + textBox3.Text + "' WHERE код_группы =" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_группы, количество_учащихся, название_специальности, группа FROM Группа, Специальность WHERE Группа.код_специальности=Специальность.код_специальности ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (this.Text == "Дисциплины")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                bool result = checkBox1.Checked;
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Дисциплины SET название_дисциплины='" + textBox2.Text + "', Экзамен=" + result + " WHERE код_дисциплины =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Дисциплины";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Дисциплины_группы")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                int idDis = getIdByDis(comboBox1.Text);
                int idGrup = getIdByGrup(comboBox2.Text);
                int idKabinet = getIdByKabinet(comboBox3.Text);
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Дисциплины_группы SET код_дисциплины=" + idDis + ", код_группы=" + idGrup + ",код_кабинета=" + idKabinet + ",количество_часов_дисциплины='" + Convert.ToInt16(textBox2.Text) + "' WHERE код_дисциплины_группы =" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_дисциплины_группы, название_дисциплины, группа, номер_кабинета, количество_часов_дисциплины FROM Дисциплины_группы, Дисциплины, Группа, Кабинет WHERE Дисциплины_группы.код_дисциплины=Дисциплины.код_дисциплины and Дисциплины_группы.код_группы=Группа.код_группы and Дисциплины_группы.код_кабинета=Кабинет.код_кабинета";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Занятие")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                int idKabinet = getIdByKabinet(comboBox1.SelectedItem.ToString());
                int idRaspis = getIdByRaspis(comboBox2.SelectedItem.ToString());
                int idDis = getIdByDis(comboBox3.SelectedItem.ToString());
                int idPrepod = getIdByPrepod(comboBox4.SelectedItem.ToString());
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Занятие SET код_кабинета=" + idKabinet + ", группа='" + textBox2.Text + "',id_расписания=" + idRaspis + ",код_дисциплины=" + idDis + ",код_преподавателя=" + idPrepod + " WHERE код_занятия =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_занятия, номер_кабинета, группа, дата_расписания,ФИО,название_дисциплины FROM Занятие, Кабинет, Расписание,Дисциплины,Преподаватель WHERE Занятие.код_кабинета=Кабинет.код_кабинета and Занятие.id_расписания=Расписание.id_расписания and Занятие.код_дисциплины=Дисциплины.код_дисциплины and Занятие.код_преподавателя=Преподаватель.код_преподавателя";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Кабинет")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Кабинет SET номер_кабинета='" + textBox2.Text + "' WHERE код_кабинета =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Кабинет";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            if (this.Text == "Преподаватель")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Преподаватель SET номер_кабинета_преподавателя='" + textBox2.Text + "',ФИО='" + textBox3.Text + "' WHERE код_преподавателя =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Преподаватель";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Расписание")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());

                int idKabinet = getIdByKabinet(comboBox1.Text);
                int idDis = getIdByDis(comboBox3.SelectedItem.ToString());
                int idGruppa = getIdByGrup(comboBox2.SelectedItem.ToString());
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Расписание SET дата_расписания='" + textBox2.Text + "',код_кабинета=" + idKabinet + " , код_группы ="+ idGruppa +", код_дисципины ="+ idDis +"  WHERE id_расписания =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectraspisan;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Специальность")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Специальность SET название_специальности='" + textBox2.Text + "' WHERE код_специальности =" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Специальность";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (this.Text == "Учащиеся")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());

                int idGrup = getIdByGrup(comboBox1.Text);
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                con.Open();
                cmd.CommandText = "UPDATE Учащиеся SET ФИО='" + textBox2.Text + "' ,телефон='" + textBox3.Text + "' ,улица='" + textBox4.Text + "' ,дата_рождения='" + Convert.ToDateTime(textBox5.Text) + "',код_группы=" + idGrup + " WHERE код_учащегося=" + ID + ""; ;
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Учащиеся.код_учащегося, Учащиеся.ФИО, Учащиеся.телефон, Учащиеся.улица, Учащиеся.дата_рождения, Группа.группа FROM Группа INNER JOIN Учащиеся ON Группа.код_группы = Учащиеся.код_группы";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public bool ClassesTeacher(bool result)
        {
            return result;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (this.Text == "Группа")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();

            }
            if (this.Text == "Дисциплины")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();

            }
            if (this.Text == "Дисциплины_группы")
            {
                comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox2.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                comboBox3.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

            }
            if (this.Text == "Занятие")
            {
                comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                comboBox2.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                comboBox3.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                comboBox4.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();

            }
            if (this.Text == "Кабинет")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();

            }
            if (this.Text == "Преподаватель")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();

            }
            if (this.Text == "Расписание")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                comboBox2.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                comboBox3.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            }
            if (this.Text == "Специальность")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();

            }
            if (this.Text == "Учащиеся")
            {

                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            }
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }

        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Сетка таблицы");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = this.Text;
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
        }
        private Application xlExcel;
        private Workbook xlWorkBook;
        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                this.QuitExcel();
                this.xlExcel = new Application { Visible = false };
                this.xlWorkBook = this.xlExcel.Workbooks.Add(Missing.Value);

                // Copy contents of grid into clipboard, open new instance of excel, a new workbook and sheet,
                // paste clipboard contents into new sheet.
                this.CopyGrid();

                var xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets.Item[1];

                try
                {
                    var cr = (Range)xlWorkSheet.Cells[1, 1];

                    try
                    {
                        cr.Select();
                        xlWorkSheet.PasteSpecial(cr, NoHTMLFormatting: true);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cr);
                    }

                    this.xlWorkBook.SaveAs(Path.Combine(Path.GetTempPath(), "ItemUpdate.xls"), XlFileFormat.xlExcel5);
                }
                finally
                {
                    Marshal.ReleaseComObject(xlWorkSheet);
                }

                MessageBox.Show("File Save Successful", "Information", MessageBoxButtons.OK);

                //// If box is checked, show the exported file. Otherwise quit Excel.
                //if (this.checkBox1.Checked)
                //{
                this.xlExcel.Visible = true;
                //}
                //else
                //{
                //    this.QuitExcel();
                //}
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            // Set the Selection Mode back to Cell Select to avoid conflict with sorting mode.
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        private void CopyGrid()
        {
            // I'm making this up...
            this.dataGridView1.SelectAll();

            var data = this.dataGridView1.GetClipboardContent();

            if (data != null)
            {
                Clipboard.SetDataObject(data, true);
            }
        }
        private void QuitExcel()
        {
            if (this.xlWorkBook != null)
            {
                try
                {
                    this.xlWorkBook.Close();
                    Marshal.ReleaseComObject(this.xlWorkBook);
                }
                catch (COMException)
                {
                }

                this.xlWorkBook = null;
            }

            if (this.xlExcel != null)
            {
                try
                {
                    this.xlExcel.Quit();
                    Marshal.ReleaseComObject(this.xlExcel);
                }
                catch (COMException)
                {
                }

                this.xlExcel = null;
            }
        }



        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {



        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form3_Shown(sender, e);
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower();
                    }
                }
            }
            if (this.Text == "Группа")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_группы, количество_учащихся, название_специальности, группа FROM Группа, Специальность WHERE Группа.код_специальности=Специальность.код_специальности  and количество_учащихся LIKE '%" + textBox1.Text.ToLower() + "%' or название_специальности LIKE '%" + textBox1.Text.ToLower() + "%' or группа LIKE '%" + textBox1.Text.ToLower() + "%'";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Дисциплины")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Дисциплины WHERE название_дисциплины LIKE '%" + textBox1.Text.ToLower() + "%' or Экзамен LIKE '%" + textBox1.Text.ToLower() + "%' ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Дисциплины_группы")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_дисциплины_группы, название_дисциплины, группа, номер_кабинета, количество_часов_дисциплины FROM Дисциплины_группы, Дисциплины, Группа, Кабинет WHERE Дисциплины_группы.код_дисциплины=Дисциплины.код_дисциплины and Дисциплины_группы.код_группы=Группа.код_группы and Дисциплины_группы.код_кабинета=Кабинет.код_кабинета and название_дисциплины LIKE '%" + textBox1.Text.ToLower() + "%' or группа LIKE '%" + textBox1.Text.ToLower() + "%' or номер_кабинета LIKE '%" + textBox1.Text.ToLower() + "%' or количество_часов_дисциплины LIKE '%" + textBox1.Text.ToLower() + "%'";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Занятие")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT код_занятия, номер_кабинета, группа, дата_расписания,ФИО,название_дисциплины FROM Занятие, Кабинет, Расписание,Дисциплины,Преподаватель WHERE Занятие.код_кабинета=Кабинет.код_кабинета and Занятие.id_расписания=Расписание.id_расписания and Занятие.код_дисциплины=Дисциплины.код_дисциплины and Занятие.код_преподавателя=Преподаватель.код_преподавателя and номер_кабинета LIKE '%" + textBox1.Text.ToLower() + "%' or группа LIKE '%" + textBox1.Text.ToLower() + "%' or дата_расписания LIKE '%" + textBox1.Text.ToLower() + "%' or ФИО LIKE '%" + textBox1.Text.ToLower() + "%'or название_дисциплины LIKE '%" + textBox1.Text.ToLower() + "%'";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Кабинет")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Кабинет WHERE номер_кабинета LIKE '%" + textBox1.Text.ToLower() + "%' ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Преподаватель")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Преподаватель WHERE номер_кабинета_преподавателя LIKE '%" + textBox1.Text.ToLower() + "%' or ФИО LIKE '%" + textBox1.Text.ToLower() + "%' or Логин LIKE '%" + textBox1.Text.ToLower() + "%'or Пароль LIKE '%" + textBox1.Text.ToLower() + "%'";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Расписание")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Расписание WHERE дата_расписания LIKE '%" + textBox1.Text.ToLower() + "%' ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Специальность")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Специальность WHERE название_специальности LIKE '%" + textBox1.Text.ToLower() + "%' ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (this.Text == "Учащиеся")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Учащиеся.код_учащегося, Учащиеся.ФИО, Учащиеся.телефон, Учащиеся.улица, Учащиеся.дата_рождения, Группа.группа FROM Группа INNER JOIN Учащиеся ON Группа.код_группы = Учащиеся.код_группы WHERE ФИО LIKE '%" + textBox1.Text.ToLower() + "%' or телефон LIKE '%" + textBox1.Text.ToLower() + "%' or улица LIKE '%" + textBox1.Text.ToLower() + "%' or дата_рождения LIKE '%" + textBox1.Text.ToLower() + "%'or группа LIKE '%" + textBox1.Text.ToLower() + "%'";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }

        }

        private void поДатеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.Text == "Расписание")
            {
                Form4 f4 = new Form4();
                f4.Show();
                loadTable(Queries.selectraspisan);
            }

        }

        private void Form3_Activated(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form3_Activated_1(object sender, EventArgs e)
        {
            if (Form4.first != null && Form4.second != null)
            {
                loadTable($@"SELECT Расписание.id_расписания, Расписание.дата_расписания, Кабинет.номер_кабинета, Группа.группа, Дисциплины.название_дисциплины
FROM Кабинет INNER JOIN (Дисциплины INNER JOIN (Группа INNER JOIN Расписание ON Группа.код_группы = Расписание.код_группы) ON Дисциплины.код_дисциплины = Расписание.код_дисципины) ON Кабинет.код_кабинета = Расписание.код_кабинета WHERE Расписание.дата_расписания between #{Form4.first}# and #{Form4.second}#");
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
    }
}









