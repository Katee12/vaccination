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
    public partial class main_add_vaccination : Form
    {
        Dictionary<string, int> studentIdMap = new Dictionary<string, int>();
        Dictionary<string, int> vaccineIdMap = new Dictionary<string, int>();
        DbConnection dbcon = new DbConnection();
        public main_add_vaccination()
        {
            InitializeComponent();
        }
        public main_add_vaccination(main main)
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void main_add_vaccination_Load(object sender, EventArgs e)
        {
            loadVaccinationCombobox();
            loadGroupCombobox();
            loadStudentsCombobox("");
        }
        private void loadVaccinationCombobox()
        {
            try
            {
                DataTable tables = dbcon.GetData("getvaccinationinfo");
                comboBox1.Items.Clear();

                foreach (DataRow row in tables.Rows)
                {
                    comboBox1.Items.Add(row["прививка"].ToString());
                    vaccineIdMap[row["прививка"].ToString()] = Convert.ToInt32(row["идентификатор_прививки"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }
        private void loadGroupCombobox()
        {
            try
            {
                DataTable tables = dbcon.GetData("getallclasses");
                comboBox2.Items.Clear();

                foreach (DataRow row in tables.Rows)
                {
                    comboBox2.Items.Add(row["группа"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }
        private void loadStudentsCombobox(string group)
        {
            try
            {
                DataTable table = null;
                comboBox3.Items.Clear();

                if (group == "")
                    table = dbcon.GetData("getstudentsinfo");
                else
                    table = dbcon.GetData("get_students_info_by_class('" + group + "')");

                if (table != null)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        comboBox3.Items.Add(row["ФИО"].ToString());
                        studentIdMap[row["ФИО"].ToString()] = Convert.ToInt32(row["идентификатор_студента"]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadStudentsCombobox(comboBox2.SelectedItem.ToString());
        }

        private void main_add_vaccination_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form main = new main(this);
            main.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string selectedVaccine = comboBox1.SelectedItem?.ToString();
                string selectedStudent = comboBox3.SelectedItem?.ToString();

                if (string.IsNullOrEmpty(selectedVaccine) || string.IsNullOrEmpty(selectedStudent))
                {
                    MessageBox.Show("Выберите прививку и студента");
                    return;
                }

                dbcon.insertIntoStudentsVaccination(
                     studentIdMap[selectedStudent],
                   vaccineIdMap[selectedVaccine],
                    dateTimePicker1.Value);

                MessageBox.Show("Прививка успешно добавлена");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных: {ex.Message}");
            }
        }
    }
}
