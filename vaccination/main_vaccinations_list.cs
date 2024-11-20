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
    public partial class main_vaccinations_list : Form
    {
        DbConnection dbcon = new DbConnection();
        public main_vaccinations_list()
        {
            InitializeComponent();
        }
        public main_vaccinations_list(main_students main_students)
        {
            InitializeComponent();
        }
        public main_vaccinations_list(main main)
        {
            InitializeComponent();
        }
        public main_vaccinations_list(main_classes main_classes)
        {
            InitializeComponent();
        }

        private void main_vaccinations_list_Load(object sender, EventArgs e)
        {
            fillDataGridView();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form main_students = new main_students(this);
            main_students.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form main = new main(this);
            main.Show();
            this.Hide();
        }

        private void main_vaccinations_list_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form main_classes = new main_classes(this);
            main_classes.Show();
            this.Hide();
        }
        private void fillDataGridView()
        {
            try
            {
                DataTable tableData = null;

                tableData = dbcon.GetData("getallvaccinationscheduledata");

                dataGridView1.DataSource = tableData;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении таблицы: {ex.Message}");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string searchText = textBox1.Text.ToLower();

                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = "";

                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter =
                    $"Convert({dataGridView1.Columns[1].Name}, 'System.String') LIKE '%{searchText}%'";

                dataGridView1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске: {ex.Message}");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Хотите удалить всю информацию об этой прививке?", "Подтверждение удаления", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    string vaccineName = "";

                    if (dataGridView1.SelectedCells.Count > 0)
                    {
                        int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                        vaccineName = dataGridView1.Rows[selectedRowIndex].Cells[1].Value.ToString();
                    }

                    if (!string.IsNullOrEmpty(vaccineName))
                    {
                        string query = $"DELETE FROM vaccination_schedule WHERE id_vaccination = (SELECT id FROM vaccination WHERE name = '{vaccineName}');";
                        dbcon.InsertDeleteUpdateData(query);

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
            else if (result == DialogResult.No)
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
                        string query = $"DELETE FROM vaccination_schedule WHERE id = @id;";
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
                        object value = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                        int id = -1;

                        string columnName = convertColumnName(dataGridView1.Columns[e.ColumnIndex].Name);

                        if (columnName == "id" || columnName == "vaccination_name")
                        {
                            MessageBox.Show("Вы не можете изменить это поле");
                            fillDataGridView();
                            return;
                        }

                        id = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value);
                        string query = null;

                        if (columnName == "necessary")
                        {
                            query = "UPDATE vaccination SET necessary = " + value + " WHERE name = '" +
                                dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + "';";

                            dbcon.InsertDeleteUpdateData(query);

                            fillDataGridView();

                            return;
                        }

                        query = "UPDATE vaccination_schedule SET " + columnName + " = @value WHERE id = " + id + ";";

                        if (columnName == "age_group")
                        {
                            switch (value)
                            {
                                case "все":
                                    value = "all";
                                    break;
                                case "дети":
                                    value = "child";
                                    break;
                                case "взрослые":
                                    value = "adult";
                                    break;
                                default:
                                    value = "error";
                                    break;
                            }
                            if (value.ToString() == "error")
                            {
                                MessageBox.Show("Некорректное значение в поле возрастной группы");
                                return;
                            }
                            query = "UPDATE vaccination_schedule SET age_group = @value::age_group WHERE id = " + id + ";";
                        }

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
                        if (selectedRow.DataGridView.Columns[cell.ColumnIndex].Name == "идентификатор_прививки")
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
                        dbcon.addVaccinationSchedule(
                            selectedRow.Cells["наименование_прививки"].Value.ToString(),
                            Convert.ToDouble(selectedRow.Cells["рекомендованный_возраст"].Value.ToString()),
                            Convert.ToInt32(selectedRow.Cells["номер_дозы"].Value),
                            Convert.ToInt32(selectedRow.Cells["максимально_доз"].Value),
                            Convert.ToInt32(selectedRow.Cells["до_следующей_прививки"].Value),
                            Convert.ToInt32(selectedRow.Cells["допустимое_отклонение"].Value),
                            selectedRow.Cells["возрастная_группа"].Value.ToString(),
                            Convert.ToBoolean(selectedRow.Cells["экстренность"].Value),
                            Convert.ToBoolean(selectedRow.Cells["обязательность"].Value));

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
                case "идентификатор_прививки":
                    convertedColumnName = "id";
                    break;
                case "наименование_прививки":
                    convertedColumnName = "vaccination_name";
                    break;
                case "рекомендованный_возраст":
                    convertedColumnName = "recommended_age";
                    break;
                case "номер_дозы":
                    convertedColumnName = "dose_number";
                    break;
                case "максимально_доз":
                    convertedColumnName = "max_doses";
                    break;
                case "до_следующей_прививки":
                    convertedColumnName = "interval_days";
                    break;
                case "допустимое_отклонение":
                    convertedColumnName = "acceptable_delay";
                    break;
                case "возрастная_группа":
                    convertedColumnName = "age_group";
                    break;
                case "экстренность":
                    convertedColumnName = "emergency";
                    break;
                case "обязательность":
                    convertedColumnName = "necessary";
                    break;
            }

            return convertedColumnName;
        }

        private void button4_Click(object sender, EventArgs e)
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
