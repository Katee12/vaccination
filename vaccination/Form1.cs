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
    public partial class Form1 : Form
    {
        DbConnection dbcon = new DbConnection();
        public Form1()
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(main main)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(main_students main_students)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(main_classes main_classes)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(main_vaccinations_list main_vaccination_list)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(admin_form admin_form)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }
        public Form1(student_info student_info)
        {
            InitializeComponent();

            this.KeyPreview = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Заполните все поля");
            }
            else
            {
                string login = textBox1.Text;
                string password = textBox2.Text;

                try
                {
                    string role = dbcon.autorization(login, password);

                    if (role == null)
                    {
                        MessageBox.Show("Ошибка в логине или пароле");
                    }
                    else if (role == "admin")
                    {
                        Form admin_form = new admin_form(this);
                        admin_form.Show();
                        this.Hide();
                    }
                    else if (role == "student")
                    {
                        Form student_info = new student_info(this, login);
                        student_info.Show();
                        this.Hide();
                    }
                    else if (role == "doctor")
                    {
                        Form main = new main(this);
                        main.Show();
                        this.Hide();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}");
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!dbcon.CheckConnection())
            {
                MessageBox.Show("Ошибка подключения к базе данных");
                Application.Exit();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Form1_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}
