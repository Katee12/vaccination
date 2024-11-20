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
    public partial class student_vaccination_history : Form
    {
        DbConnection dbcon = new DbConnection();
        string login;
        public student_vaccination_history()
        {
            InitializeComponent();
        }
        public student_vaccination_history(student_info student_info, string login)
        {
            this.login = login;
            InitializeComponent();
        }

        private void student_vaccination_history_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form student_info = new student_info(this, login);
            student_info.Show();
            this.Hide();
        }

        private void student_vaccination_history_Load(object sender, EventArgs e)
        {
            setVaccinationsCombobox();
            loadDataGridView();
        }
        private void loadDataGridView()
        {
            try
            {
                DataTable tableData = null;
                tableData = dbcon.GetData("get_vaccination_history_by_login('" + login + "');");
                dataGridView1.DataSource = tableData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }
        private void setVaccinationsCombobox()
        {
            try
            {
                DataTable tables = dbcon.addNullValue(dbcon.GetData("getallvaccinationsname"), "название_прививки");
                comboBox1.Items.Clear();

                foreach (DataRow row in tables.Rows)
                {
                    comboBox1.Items.Add(row["название_прививки"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }
        private void ApplyFilters()
        {
            string selectedVaccine = comboBox1.SelectedItem?.ToString();
            string date = textBox1.Text;

            List<string> filters = new List<string>();

            if (!string.IsNullOrEmpty(selectedVaccine) && selectedVaccine != "не задано")
            {
                filters.Add($"Convert([прививка], 'System.String') LIKE '%{selectedVaccine}%'");
            }

            if (!string.IsNullOrEmpty(date))
            {
                filters.Add($"Convert([дата], 'System.String') LIKE '%{date}%'");
            }

            string filterString = string.Join(" AND ", filters);

            if (!string.IsNullOrEmpty(filterString))
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = filterString;
            }
            else
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = "";
            }

            dataGridView1.Refresh();
        }
    }
}
