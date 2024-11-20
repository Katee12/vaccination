using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vaccination
{
    public partial class student_info : Form
    {
        DbConnection dbcon = new DbConnection();
        string login;
        int id = -1;
        bool isNameChanged = false;
        bool isLastnameChanged = false;
        bool isPatronymicChanged = false;
        bool isBirthdayChanged = false;
        public student_info()
        {
            InitializeComponent();
        }
        public student_info(Form1 form1, string login)
        {
            this.login = login;
            InitializeComponent();
        }
        public student_info(student_vaccination_history student_vaccination_history, string login)
        {
            this.login = login;
            InitializeComponent();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void student_info_Load(object sender, EventArgs e)
        {
            setTextBoxex();
        }
        private void setTextBoxex()
        {
            try
            {
                DataTable studentinto = null;
                studentinto = dbcon.GetData("get_students_info_by_login('" + login + "');");

                if (studentinto.Rows.Count > 0)
                {
                    DataRow row = studentinto.Rows[0];

                    id = Convert.ToInt32(row["идентификатор_студента"]);
                    textBox1.Text = row["имя"].ToString();
                    textBox2.Text = row["фамилия"].ToString();
                    textBox3.Text = row["отчество"].ToString();
                    maskedTextBox1.Text = row["дата_рождения"].ToString();
                    textBox5.Text = row["номер_зачетки"].ToString();
                    textBox6.Text = row["специальность"].ToString();
                    textBox7.Text = row["группа"].ToString();
                }

                isNameChanged = false;
                isLastnameChanged = false;
                isPatronymicChanged = false;
                isBirthdayChanged = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении данных: {ex.Message}");
            }
        }

        private void student_info_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            isNameChanged = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы точно хотите изменить данные?", "Подтверждение изменения", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    string query = null;

                    if (isNameChanged)
                    {
                        query = "UPDATE students SET name = @value WHERE id = " + id;

                        dbcon.UpdateStudentsData(query, textBox1.Text.ToString());
                    }

                    if (isLastnameChanged)
                    {
                        query = "UPDATE students SET lastname = @value WHERE id = " + id;

                        dbcon.UpdateStudentsData(query, textBox2.Text.ToString());
                    }

                    if (isPatronymicChanged)
                    {
                        query = "UPDATE students SET patronymic = @value WHERE id = " + id;

                        dbcon.UpdateStudentsData(query, textBox3.Text.ToString());
                    }

                    if (isBirthdayChanged)
                    {
                        query = "UPDATE students SET birthday = @value WHERE id = " + id;

                        dbcon.UpdateStudentsData(query,
                            dbcon.getNewValue("DateTime", maskedTextBox1.Text.ToString()));
                    }

                    MessageBox.Show("Данные успешно обновлены");

                    setTextBoxex();

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при заполнении данных: {ex.Message}");
                    setTextBoxex();
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            isLastnameChanged = true;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            isPatronymicChanged = true;
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            
        }

        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            isBirthdayChanged = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form form1 = new Form1(this);
            form1.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form student_vaccination_history = new student_vaccination_history(this, login);
            student_vaccination_history.Show();
            this.Hide();
        }
    }
}
