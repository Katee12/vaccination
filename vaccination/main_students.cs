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
    public partial class main_students : Form
    {
        DbConnection dbcon = new DbConnection();
        public main_students()
        {
            InitializeComponent();
        }
        public main_students(main main)
        {
            InitializeComponent();
        }
        public main_students(main_classes main_classes)
        {
            InitializeComponent();
        }
        public main_students(main_vaccinations_list main_vaccinations_list)
        {
            InitializeComponent();
        }

        private void main_students_Load(object sender, EventArgs e)
        {
            setSpecializationCombobox();
            setGroupCombobox();
            fillDataGridView();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form main = new main(this);
            main.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form main_vaccinations_list = new main_vaccinations_list(this);
            main_vaccinations_list.Show();
            this.Hide();
        }

        private void main_students_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void fillDataGridView()
        {
            try
            {
                DataTable tableData = null;

                tableData = dbcon.GetData("getallstudentsdata");

                dataGridView1.DataSource = tableData;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении таблицы: {ex.Message}");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }
        private void setSpecializationCombobox()
        {
            try
            {
                DataTable tables = dbcon.addNullValue(dbcon.GetData("getallspecialization"), "специальность");
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void button4_Click(object sender, EventArgs e)
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
                        string query = $"DELETE FROM students WHERE id = @id;";
                        dbcon.DeleteData(query, id);

                        fillDataGridView();
                    }
                    else
                    {
                        MessageBox.Show("Невозможно удалить строку");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении данных: {ex.Message}");
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int selectedRowIndex = dataGridView1.CurrentRow.Index + 1;
                int rowCount = getRowCount();

                if (selectedRowIndex <= rowCount)
                {
                    DialogResult result = MessageBox.Show("Вы точно хотите изменить данные?", "Подтверждение изменения", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        string typeColumn = dataGridView1.Columns[e.ColumnIndex].ValueType.Name;

                        object value = dbcon.getNewValue(typeColumn, dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());

                        string columnName = convertColumnName(dataGridView1.Columns[e.ColumnIndex].Name);

                        int id = -1;

                        if (columnName == "id" || columnName == "specialization")
                        {
                            MessageBox.Show("Вы не можете изменить это поле");
                            fillDataGridView();
                            return;
                        }

                        id = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value);
                        string query = null;

                        if (columnName == "class")
                        {
                            int id_classes = dbcon.getID(columnName, value);

                            if (id_classes == -1)
                            {
                                DialogResult result2 = MessageBox.Show("Такой группы/специальности нет в базе. Хотите добавить?", "Ошибка", MessageBoxButtons.YesNo, MessageBoxIcon.Error);

                                if (result == DialogResult.Yes)
                                {
                                    Form main_classes = new main_classes(this);
                                    main_classes.Show();
                                    this.Hide();
                                    return;
                                }
                            }
                            else
                            {
                                query = "UPDATE students SET id_classes = (SELECT id FROM classes WHERE "
                                + columnName + " = '" + value + "') WHERE id = " + id + ";";
                                dbcon.InsertDeleteUpdateData(query);

                                fillDataGridView();
                            }
                        }
                        else
                        {
                            query = "UPDATE students SET " + columnName + " = @value WHERE id = " + id + ";";
                            dbcon.UpdateData(query, value);

                            fillDataGridView();
                        }
                    }
                }
                else
                {
                    DataGridViewRow selectedRow = dataGridView1.CurrentRow;

                    bool allValuesFilled = true;
                    foreach (DataGridViewCell cell in selectedRow.Cells)
                    {
                        if (selectedRow.DataGridView.Columns[cell.ColumnIndex].Name == "идентификатор_студента")
                        {
                            continue;
                        }

                        if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        {
                            allValuesFilled = false;
                            break;
                        }
                    }

                    if (allValuesFilled)
                    {
                        dbcon.addStudents(
                            selectedRow.Cells["фамилия"].Value.ToString(),
                            selectedRow.Cells["имя"].Value.ToString(),
                            selectedRow.Cells["отчество"].Value.ToString(),
                            selectedRow.Cells["номер_зачетки"].Value,
                            dbcon.getNewValue(
                                 selectedRow.Cells["дата_рождения"].ValueType.Name, selectedRow.Cells["дата_рождения"].Value.ToString()),
                             selectedRow.Cells["специальность"].Value.ToString(),
                             selectedRow.Cells["группа"].Value.ToString());

                        fillDataGridView();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении/добавлении данных: {ex.Message}");
                fillDataGridView();
            }
        }
    
        private string convertColumnName(string columnName)
        {
            string convertedColumnName = null;

            switch (columnName)
            {
                case "идентификатор_студента":
                    convertedColumnName = "id";
                    break;
                case "фамилия":
                    convertedColumnName = "lastname";
                    break;
                case "имя":
                    convertedColumnName = "name";
                    break;
                case "отчество":
                    convertedColumnName = "patronymic";
                    break;
                case "дата_рождения":
                    convertedColumnName = "birthday";
                    break;
                case "номер_зачетки":
                    convertedColumnName = "record_book_number";
                    break;
                case "группа":
                    convertedColumnName = "class";
                    break;
                case "специальность":
                    convertedColumnName = "specialization";
                    break;
            }

            return convertedColumnName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form main_classes = new main_classes(this);
            main_classes.Show();
            this.Hide();
        }
        private int getRowCount()
        {
            int rowCount = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                bool allValuesFilled = true;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    {
                        allValuesFilled = false;
                        break;
                    }
                }

                if (allValuesFilled)
                {
                    rowCount++;
                }
            }

            return rowCount;
        }
        private void ApplyFilters()
        {
            string selectedSpecialty = comboBox1.SelectedItem?.ToString();
            string selectedGroup = comboBox2.SelectedItem?.ToString();
            string recordBook = textBox1.Text.ToLower();
            string lastName = textBox2.Text.ToLower();

            List<string> filters = new List<string>();

            if (!string.IsNullOrEmpty(selectedSpecialty) && selectedSpecialty != "не задано")
            {
                filters.Add($"Convert([специальность], 'System.String') LIKE '%{selectedSpecialty}%'");
            }

            if (!string.IsNullOrEmpty(selectedGroup) && selectedGroup != "не задано")
            {
                filters.Add($"Convert([группа], 'System.String') LIKE '%{selectedGroup}%'");
            }

            if (!string.IsNullOrEmpty(recordBook))
            {
                filters.Add($"Convert([номер_зачетки], 'System.String') LIKE '%{recordBook}%'");
            }

            if (!string.IsNullOrEmpty(lastName))
            {
                filters.Add($"Convert([фамилия], 'System.String') LIKE '%{lastName}%'");
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
    }
}
