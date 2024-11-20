using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vaccination
{
    class DbConnection
    {
        private string connectionString;
        public DbConnection()
        {
            string host = "172.20.7.8";
            string port = "5432";
            string username = "st1991";
            string password = "pwd1991";
            string database = "vaccination";

            connectionString = $"Host={host};Port={port};Username={username};Password={password};Database={database}";
        }
        public bool CheckConnection()
        {
            try
            {
                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    connection.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public string autorization(string login, string password)
        {
            string roleName = null;

            string query = "SELECT roles FROM users WHERE login = @login AND password = md5(@password)";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("login", login);
                    command.Parameters.AddWithValue("password", password);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            roleName = reader.GetString(0);
                        }
                    }
                }
            }

            return roleName;
        }
        public DataTable GetData(string tablename)
        {
            DataTable table = new DataTable();

            string query = $"SELECT * FROM " + tablename;

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(command);
                    adapter.Fill(table);
                }
            }

            return table;
        }

        public DataTable addNullValue(DataTable table, string columnName)
        {
            if (table.Columns.Contains(columnName))
            {
                DataRow newRow = table.NewRow();
                newRow[columnName] = "не задано"; 
                table.Rows.InsertAt(newRow, 0);
            }

            return table;
        }
        public void DeleteData(string query, int id)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("id", id);

                    command.ExecuteNonQuery();
                }
            }
        }
        public object getNewValue(string typeColumn, string value)
        {
            object newValue = null;

            switch (typeColumn)
            {
                case "DateTime":
                    newValue = Convert.ToDateTime(value);
                    break;
                case "Decimal":
                    newValue = Convert.ToDecimal(value);
                    break;
                case "String":
                    newValue = value;
                    break;
                case "Int32":
                    newValue = Convert.ToInt32(value);
                    break;
                case "Boolean":
                    newValue = Convert.ToBoolean(value);
                    break;
            }

            return newValue;
        }
        public void UpdateData(string query, object parameterValue)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("value", parameterValue);
                    command.ExecuteNonQuery();
                }
            }
        }
        public int getID(string columnName, object value)
        {
            int id = -1;

            string query = "select id from classes where " +
                columnName + " = '" +
                value + "'";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id = reader.GetInt32(0);
                        }
                    }
                }
            }

            return id;
        }
        public void InsertDeleteUpdateData(string query)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
        public bool getGroup(string value)
        {
            bool isGroupExist = true;

            string query = "SELECT EXISTS ( SELECT 1 FROM classes c where c.class = '" +
                value + "')";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            isGroupExist = reader.GetBoolean(0);
                        }
                    }
                }
            }

            return isGroupExist;
        }
        public bool getSpecialization(string value)
        {
            bool isSpecializaionExist = true;

            string query = "SELECT EXISTS ( SELECT 1 FROM pg_enum " +
                "WHERE enumtypid = 'specialization'::regtype " +
                "AND enumlabel = '" +
                value + "');";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            isSpecializaionExist = reader.GetBoolean(0);
                        }
                    }
                }
            }

            return isSpecializaionExist;
        }
        public void addStudents(object lastname, object name, object patronymic, object record_book_number, object birthday, object specialization, object group)
        {
            if (getRecordBookNumber(record_book_number))
            {
                MessageBox.Show("Номер зачетной книжки не должен повторяться");
                return;
            }

            int id_classes = getClassesIDBySG(specialization, group);

            if (id_classes == -1)
            {
                MessageBox.Show("Такой группы на такой специальности нет");
                return;
            }
               
            string query = "INSERT INTO students (lastname, name, patronymic, record_book_number, birthday, id_classes)" +
                               " VALUES (@lastname, @name, @patronymic, @record_book_number, @birthday, @id_classes);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("lastname", lastname);
                    command.Parameters.AddWithValue("name", name);
                    command.Parameters.AddWithValue("patronymic", patronymic);
                    command.Parameters.AddWithValue("record_book_number", record_book_number);
                    command.Parameters.AddWithValue("birthday", birthday);
                    command.Parameters.AddWithValue("id_classes", id_classes);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Данные успешно добавлены");
                }
            }
        }
        public int getClassesIDBySG(object specialization, object group)
        {
            int id = -1;

            string query = "SELECT id FROM classes WHERE specialization = @specialization::specialization AND class = @group;";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("specialization", specialization);
                    command.Parameters.AddWithValue("group", group);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id = reader.GetInt32(0);
                        }
                    }
                }
            }

            return id;
        }
        public bool getRecordBookNumber(object value)
        {
            bool isBookExist = true;

            string query = "SELECT EXISTS ( SELECT 1 FROM students " +
                "WHERE record_book_number = @value)";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("value", value);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            isBookExist = reader.GetBoolean(0);
                        }
                    }
                }
            }

            return isBookExist;
        }
        public void insertIntoStudentsVaccination(int id_students, int id_vaccination_schedule, DateTime vaccination_date)
        {
            string query = "INSERT INTO students_vaccination (id_students, id_vaccination_schedule, vaccination_date) VALUES (" +
                "@id_students, @id_vaccination_schedule, @vaccination_date);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("id_students", id_students);
                    command.Parameters.AddWithValue("id_vaccination_schedule", id_vaccination_schedule);
                    command.Parameters.AddWithValue("vaccination_date", vaccination_date);

                    command.ExecuteNonQuery();
                }
            }
        }
        public void addVaccinationSchedule(
            string vaccination_name, 
            double recommended_age, 
            int dose_number, 
            int max_doses, 
            int interval_days, 
            int acceptable_delay, 
            string age_group,
            bool emergency,
            bool necessary)
        {
            
            if (getVaccinationIDByName(vaccination_name) == -1)
                InsertDeleteUpdateData("INSERT INTO vaccination (name, necessary) VALUES ('" + 
                    vaccination_name + "', '" + 
                    necessary + "');");

            int id_vaccination = getVaccinationIDByName(vaccination_name);

            if (max_doses < dose_number)
            {
                MessageBox.Show("Некорректные значения в дозах");
                return;
            }
            
            switch (age_group)
            {
                case "все":
                    age_group = "all";
                    break;
                case "дети":
                    age_group = "child";
                    break;
                case "взрослые":
                    age_group = "adult";
                    break;
                default:
                    age_group = "error";
                    break;
            }

            if (age_group == "error")
            {
                MessageBox.Show("Некорректное значение в поле возрастной группы");
                return;
            }

            string query = "INSERT INTO vaccination_schedule (id_vaccination, recommended_age, dose_number, max_doses, interval_days, acceptable_delay, age_group, emergency)" +
                               " VALUES (@id_vaccination, @recommended_age, @dose_number, @max_doses, @interval_days, @acceptable_delay, @age_group::age_group, @emergency);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("id_vaccination", id_vaccination);
                    command.Parameters.AddWithValue("recommended_age", recommended_age);
                    command.Parameters.AddWithValue("dose_number", dose_number);
                    command.Parameters.AddWithValue("max_doses", max_doses);
                    command.Parameters.AddWithValue("interval_days", interval_days);
                    command.Parameters.AddWithValue("acceptable_delay", acceptable_delay);
                    command.Parameters.AddWithValue("age_group", age_group);
                    command.Parameters.AddWithValue("emergency", emergency);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Данные успешно добавлены");
                }
            }
        }
        public int getVaccinationIDByName(object vaccination_name)
        {
            int id = -1;

            string query = "SELECT id FROM vaccination WHERE name = @vaccination_name;";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("vaccination_name", vaccination_name);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id = reader.GetInt32(0);
                        }
                    }
                }
            }

            return id;
        }
        public DataTable GetAllTables()
        {
            DataTable dt = new DataTable();
            string query = "SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'  AND table_type = 'BASE TABLE';";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand(query, connection))
                {
                    using (var adapter = new NpgsqlDataAdapter(command))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }
        public void InsertIntoUsers(string login, string password, string roles)
        {
            if (roles == "admin" || roles == "doctor")
            {
                MessageBox.Show("Такие пользователи уже существуют");
                return;
            }

            string query = "INSERT INTO users (login, password, roles) VALUES (@login, md5(@password), @roles::roles);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("login", login);
                    command.Parameters.AddWithValue("password", password);
                    command.Parameters.AddWithValue("roles", roles);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void InsertIntoStudents(
            string name, 
            string lastname, 
            string patronymic, 
            object birthday, 
            int record_book_number, 
            int id_users, 
            int id_classes)
        {
            if (getRecordBookNumber(record_book_number))
            {
                MessageBox.Show("Номер зачетной книжки не должен повторяться");
                return;
            }


            string query = "INSERT INTO students (name, lastname, patronymic, birthday, record_book_number, id_users, id_classes)" +
                               " VALUES (@name, @lastname, @patronymic, @birthday, @record_book_number, @id_users, @id_classes);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("name", name);
                    command.Parameters.AddWithValue("lastname", lastname);
                    command.Parameters.AddWithValue("patronymic", patronymic);
                    command.Parameters.AddWithValue("birthday", birthday);
                    command.Parameters.AddWithValue("record_book_number", record_book_number);
                    command.Parameters.AddWithValue("id_users", id_users);
                    command.Parameters.AddWithValue("id_classes", id_classes);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void InsertIntoClasses(string specialization, string group)
        {
            if (!getSpecialization(specialization))
                InsertDeleteUpdateData("ALTER TYPE specialization ADD VALUE '" + specialization + "';");

            string query = "INSERT INTO classes (specialization, class)" +
                               " VALUES (@specialization::specialization, @class);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("specialization", specialization);
                    command.Parameters.AddWithValue("class", group);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void InsertIntoVaccination(string name, bool necessary)
        {
            string query = "INSERT INTO vaccination (name, necessary)" +
                               " VALUES (@name, @necessary);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("name", name);
                    command.Parameters.AddWithValue("necessary", necessary);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void InsertIntoVaccinationSchedule(
            int id_vaccination,
            double recommended_age,
            int dose_number,
            int max_doses,
            int interval_days,
            int acceptable_delay,
            bool emergency, 
            string age_group)
        {
            if (max_doses < dose_number)
            {
                MessageBox.Show("Некорректные значения в дозах");
                return;
            }

            string query = "INSERT INTO vaccination_schedule (id_vaccination, recommended_age, dose_number, max_doses, interval_days, acceptable_delay, emergency, age_group)" +
                               " VALUES (@id_vaccination, @recommended_age, @dose_number, @max_doses, @interval_days, @acceptable_delay, @emergency, @age_group::age_group);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("id_vaccination", id_vaccination);
                    command.Parameters.AddWithValue("recommended_age", recommended_age);
                    command.Parameters.AddWithValue("dose_number", dose_number);
                    command.Parameters.AddWithValue("max_doses", max_doses);
                    command.Parameters.AddWithValue("interval_days", interval_days);
                    command.Parameters.AddWithValue("acceptable_delay", acceptable_delay);
                    command.Parameters.AddWithValue("emergency", emergency);
                    command.Parameters.AddWithValue("age_group", age_group);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void InsertIntoStudentsVaccination(
            int id_students,
            int id_vaccination_schedule,
            object vaccination_date)
        {
            string query = "INSERT INTO students_vaccination (id_students, id_vaccination_schedule, vaccination_date)" +
                               " VALUES (@id_students, @id_vaccination_schedule, @vaccination_date);";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("id_students", id_students);
                    command.Parameters.AddWithValue("id_vaccination_schedule", id_vaccination_schedule);
                    command.Parameters.AddWithValue("vaccination_date", vaccination_date);
                    command.ExecuteNonQuery();
                }
            }
        }
        public void UpdateStudentsData(string query, object parameterValue)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("value", parameterValue);
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
