using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vaccination
{
    public partial class main : Form
    {
        DbConnection dbcon = new DbConnection();
        public main()
        {
            InitializeComponent();
        }
        public main(Form1 form1)
        {
            InitializeComponent();
        }
        public main(main_students main_students)
        {
            InitializeComponent();
        }
        public main(main_classes main_classes)
        {
            InitializeComponent();
        }
        public main(main_vaccinations_list main_vaccinations_list)
        {
            InitializeComponent();
        }
        public main(main_add_vaccination main_add_vaccination)
        {
            InitializeComponent();
        }

        private void main_Load(object sender, EventArgs e)
        {
            setVaccinationsCombobox();
            setGroupCombobox();
            fillDataGridView();
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

        private void main_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
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
        private void setGroupCombobox()
        {
            try
            {
                DataTable tables = dbcon.addNullValue(dbcon.GetData("getallclasses"), "группа");
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
        private void fillDataGridView()
        {
            try
            {
                DataTable tableData = null;
                tableData = dbcon.GetData("getallvaccination");
                dataGridView1.DataSource = tableData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении таблицы: {ex.Message}");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = true;

                    Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();

                    int rows = dataGridView1.Rows.Count;
                    int cols = dataGridView1.Columns.Count;
                    Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(doc.Range(), rows + 1, cols);

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        table.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            var cellValue = dataGridView1.Rows[i].Cells[j].Value;

                            table.Cell(i + 2, j + 1).Range.Text = cellValue != null ? cellValue.ToString() : string.Empty;
                        }
                    }

                    table.Borders.Enable = 1;

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string fileName = saveFileDialog1.FileName;

                        if (Path.GetExtension(fileName).ToLower() != ".docx")
                        {
                            MessageBox.Show("Пожалуйста, выберите файл с расширением .docx.");
                            return;
                        }

                        object fileNameObj = fileName;
                        doc.SaveAs2(ref fileNameObj);

                        MessageBox.Show("Таблица сохранена в Word.");
                    }
                }
                else
                {
                    MessageBox.Show("Нет данных для сохранения.");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении таблицы: {ex.Message}");
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы точно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

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
                        string query = $"DELETE FROM students_vaccination WHERE id = @id;";
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

        private void button7_Click(object sender, EventArgs e)
        {
            Form main_classes = new main_classes(this);
            main_classes.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form main_add_vaccination = new main_add_vaccination(this);
            main_add_vaccination.Show();
            this.Hide();
        }
        private void ApplyFilters()
        {
            string selectedVaccine = comboBox1.SelectedItem?.ToString();
            string selectedGroup = comboBox2.SelectedItem?.ToString();
            string students = textBox1.Text.ToLower();

            List<string> filters = new List<string>();

            if (!string.IsNullOrEmpty(selectedVaccine) && selectedVaccine != "не задано")
            {
                filters.Add($"Convert([наименование_прививки], 'System.String') LIKE '%{selectedVaccine}%'");
            }

            if (!string.IsNullOrEmpty(selectedGroup) && selectedGroup != "не задано")
            {
                filters.Add($"Convert([группа], 'System.String') LIKE '%{selectedGroup}%'");
            }

            if (!string.IsNullOrEmpty(students))
            {
                filters.Add($"Convert([фио_студента], 'System.String') LIKE '%{students}%'");
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

        private void button8_Click(object sender, EventArgs e)
        {
            Form form1 = new Form1(this);
            form1.Show();
            this.Hide();
        }
    }
}
