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
    public partial class admin_form : Form
    {
        DbConnection dbcon = new DbConnection();
        public admin_form()
        {
            InitializeComponent();
            comboboxLoad();
        }
        public admin_form(Form1 form1)
        {
            InitializeComponent();
            comboboxLoad();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form form1 = new Form1(this);
            form1.Show();
            this.Hide();
        }
        private void comboboxLoad()
        {
            try
            {
                DataTable tables = dbcon.GetAllTables();
                comboBox1.Items.Clear();

                foreach (DataRow row in tables.Rows)
                {
                    comboBox1.Items.Add(row["table_name"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке таблиц: {ex.Message}");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillDataGridView();
        }
        private void fillDataGridView()
        {
            try
            {
                string selectedTable = comboBox1.SelectedItem.ToString();
                DataTable tableData = null;

                tableData = dbcon.GetData(selectedTable);

                dataGridView1.DataSource = tableData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении таблицы: {ex.Message}");
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int selectedRowIndex = dataGridView1.CurrentRow.Index + 1;
                int rowCount = getRowCount();
                string tableName = comboBox1.SelectedItem.ToString();

                if (selectedRowIndex <= rowCount)
                {
                    DialogResult result = MessageBox.Show("Вы точно хотите изменить данные?", "Подтверждение изменения", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        string typeColumn = dataGridView1.Columns[e.ColumnIndex].ValueType.Name;

                        object value = dbcon.getNewValue(typeColumn, dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());

                        string columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                        int id = -1;

                        if (columnName == "id")
                        {
                            MessageBox.Show("Вы не можете изменить это поле");
                            fillDataGridView();
                            return;
                        }

                        id = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value);

                        string query = null;

                        if (columnName == "password")
                            query = "UPDATE " + tableName + " SET " + columnName + " = MD5(@value) WHERE id = " + id + ";";
                        else if (columnName == "specialization")
                        {
                            if (!dbcon.getSpecialization(value.ToString()))
                                dbcon.InsertDeleteUpdateData("ALTER TYPE specialization ADD VALUE '" + value + "';");
                            
                            query = "UPDATE " + tableName + " SET " + columnName + " = @value::specialization WHERE id = " + id + ";";
                        }
                        else if (columnName == "roles")
                        {
                            if (value.ToString() == "admin" || value.ToString() == "doctor")
                            {
                                MessageBox.Show("Вы не можете установить такое значение");
                                return;
                            }
                            query = "UPDATE " + tableName + " SET " + columnName + " = @value WHERE id = " + id + ";";
                        }
                        else
                            query = "UPDATE " + tableName + " SET " + columnName + " = @value WHERE id = " + id + ";";

                        dbcon.UpdateData(query, value);

                        fillDataGridView();
                    }
                }
                else
                {
                    DataGridViewRow selectedRow = dataGridView1.CurrentRow;

                    bool allValuesFilled = true;
                    foreach (DataGridViewCell cell in selectedRow.Cells)
                    {
                        if (selectedRow.DataGridView.Columns[cell.ColumnIndex].Name == "id" 
                            || selectedRow.DataGridView.Columns[cell.ColumnIndex].Name == "next_vaccination_date"
                            || selectedRow.DataGridView.Columns[cell.ColumnIndex].Name == "status")
                        {
                            continue;
                        }

                        if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        {
                            allValuesFilled = false;
                            break;
                        }
                    }

                    if (!allValuesFilled)
                        return;

                    switch (tableName)
                    {
                        case "users":
                            dbcon.InsertIntoUsers(
                             selectedRow.Cells["login"].Value.ToString(),
                             selectedRow.Cells["password"].Value.ToString(),
                             selectedRow.Cells["roles"].Value.ToString());
                            break;
                        case "students":
                            dbcon.InsertIntoStudents(
                             selectedRow.Cells["name"].Value.ToString(),
                             selectedRow.Cells["lastname"].Value.ToString(),
                             selectedRow.Cells["patronymic"].Value.ToString(),
                             dbcon.getNewValue(selectedRow.Cells["birthday"].ValueType.Name, selectedRow.Cells["birthday"].Value.ToString()),
                             Convert.ToInt32(selectedRow.Cells["record_book_number"].Value),
                             Convert.ToInt32(selectedRow.Cells["id_users"].Value),
                             Convert.ToInt32(selectedRow.Cells["id_classes"].Value));
                            break;
                        case "classes":
                            dbcon.InsertIntoClasses(
                                selectedRow.Cells["specialization"].Value.ToString(),
                                selectedRow.Cells["class"].Value.ToString());
                            break;
                        case "vaccination":
                            dbcon.InsertIntoVaccination(
                                selectedRow.Cells["name"].Value.ToString(),
                                Convert.ToBoolean(selectedRow.Cells["necessary"].Value));
                            break;
                        case "vaccination_schedule":
                            dbcon.InsertIntoVaccinationSchedule(
                                Convert.ToInt32(selectedRow.Cells["id_vaccination"].Value),
                                Convert.ToDouble(selectedRow.Cells["recommended_age"].Value),
                                Convert.ToInt32(selectedRow.Cells["dose_number"].Value),
                                Convert.ToInt32(selectedRow.Cells["max_doses"].Value),
                                Convert.ToInt32(selectedRow.Cells["interval_days"].Value),
                                Convert.ToInt32(selectedRow.Cells["acceptable_delay"].Value),
                                Convert.ToBoolean(selectedRow.Cells["emergency"].Value),
                                selectedRow.Cells["age_group"].Value.ToString());
                            break;
                        case "students_vaccination":
                            dbcon.InsertIntoStudentsVaccination(
                                Convert.ToInt32(selectedRow.Cells["id_students"].Value),
                                Convert.ToInt32(selectedRow.Cells["id_vaccination_schedule"].Value),
                                dbcon.getNewValue(selectedRow.Cells["vaccination_date"].ValueType.Name, selectedRow.Cells["vaccination_date"].Value.ToString()));
                            break;
                    }
                    fillDataGridView();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении/добавлении данных: {ex.Message}");
                fillDataGridView();
            }
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

        private void admin_form_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы точно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    string tableName = comboBox1.SelectedItem?.ToString();

                    int id = -1;

                    if (dataGridView1.SelectedCells.Count > 0)
                    {
                        int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                        id = Convert.ToInt32(dataGridView1.Rows[selectedRowIndex].Cells[0].Value);
                    }

                    if (id != -1 && !string.IsNullOrEmpty(tableName))
                    {
                        string query = $"DELETE FROM " + tableName + " WHERE id = @id;";
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string searchText = textBox1.Text.ToLower();

                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = "";

                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"Convert({dataGridView1.Columns[1].Name}, 'System.String') LIKE '%{searchText}%'";

                dataGridView1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске: {ex.Message}");
            }
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
