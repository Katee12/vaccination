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
    public partial class main_classes : Form
    {
        DbConnection dbcon = new DbConnection();
        public main_classes()
        {
            InitializeComponent();
        }
        public main_classes(main main)
        {
            InitializeComponent();
        }
        public main_classes(main_students main_students)
        {
            InitializeComponent();
        }
        public main_classes(main_vaccinations_list main_vaccinations_list)
        {
            InitializeComponent();
        }

        private void main_classes_Load(object sender, EventArgs e)
        {
            setSpecializationCombobox();
            fillDataGridView();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form main = new main(this);
            main.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form main_students = new main_students(this);
            main_students.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form main_vaccinations_list = new main_vaccinations_list(this);
            main_vaccinations_list.Show();
            this.Hide();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void setSpecializationCombobox()
        {
            try
            {
                DataTable tables = dbcon.GetData("getallspecialization");
                comboBox1.Items.Clear();

                foreach (DataRow row in tables.Rows)
                {
                    comboBox1.Items.Add(row["специальность"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }
        private void fillDataGridView()
        {
            try
            {
                DataTable tableData = null;

                tableData = dbcon.GetData("getdatafromclasses");

                dataGridView1.DataSource = tableData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении таблицы: {ex.Message}");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы точно хотите удалить группу?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    int id = -1;

                    if (dataGridView1.SelectedCells.Count > 0)
                    {
                        int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                        id = Convert.ToInt32(dataGridView1.Rows[selectedRowIndex].Cells[0].Value);
                    }

                    if (id != -1)
                    {
                        string query = $"DELETE FROM classes WHERE id = @id;";
                        dbcon.DeleteData(query, id);

                        fillDataGridView();
                    }
                    else
                    {
                        MessageBox.Show("Выберите строку для удаления.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении данных: {ex.Message}");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedItem is null || textBox1.Text == "")
                {
                    MessageBox.Show("Заполните все поля");
                    return;
                }
                string specialization = comboBox1.SelectedItem.ToString();
                string new_class = textBox1.Text;

                if (dbcon.getGroup(new_class))
                {
                    MessageBox.Show("Такая группа уже существует");
                    return;
                }

                string query = "INSERT INTO classes (class, specialization) VALUES ('" +
                new_class + "', '" +
                specialization + "');";
                dbcon.InsertDeleteUpdateData(query);
                MessageBox.Show("Группа успешно добавлена");
                fillDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных: {ex.Message}");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox2.Text == "")
                {
                    MessageBox.Show("Заполните поле новым значением");
                    return;
                }

                DialogResult result = MessageBox.Show("Вы точно хотите добавить новую специальность?", "Подтверждение добавления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string new_specialization = textBox2.Text;

                    if (dbcon.getSpecialization(new_specialization))
                    {
                        MessageBox.Show("Такая специальность уже существует");
                        return;
                    }

                    string query = "ALTER TYPE specialization ADD VALUE '" +
                    new_specialization + "';";
                    MessageBox.Show("Новая специальность успешно добавлена");
                    setSpecializationCombobox();
                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении данных: {ex.Message}");
            }
        }

        private void main_classes_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form form1 = new Form1(this);
            form1.Show();
            this.Hide();
        }
    }
}
