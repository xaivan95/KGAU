using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace KGAU
{
    public partial class Main_users : Form
    {
        List<Data> combo_pom = new List<Data>();
        DataTable pomes = new DataTable();
        List<DataRow> uslugi = new List<DataRow>();
        List<DataRow> oborud = new List<DataRow>();
        int _id;
        string title;
        string texts;
        public static string desktopPath = Directory.GetCurrentDirectory();//путь к exe программы
        Object _missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        Word._Application application;
        Word._Document document;
        string connectionString = "server=localhost;user=root;database=ritualdb;password=2244;";


        //Сохранение отчетов
        Excel.Application app;
        Excel.Workbook workBook;
        Excel.Worksheet sheet;



        public Main_users(int id)
        {
            InitializeComponent();
            _id = id;
            this.Text = "Окно пользователя - просмотр заявок";
            title = this.Text;
            texts = "Вы находитесь в окне просмотра Ваших заявок. Для получения иформации по заказу - нажмите на него. Для получения информаци о входящей в состав заказа услуги или оборудовании наведите на него курсор мыши";
            saveFileDialog1.Filter = "Word document|*.doc";//формат выходных данных
            saveFileDialog1.Title = "Save the Word Document";

            saveFileDialog2.Filter = "Excel document|*.xls";//формат выходных данных
            saveFileDialog2.Title = "Save the Excel Document";
        }

        private void Main_users_Load(object sender, EventArgs e)
        {
            Load_oboryd();
            Load_uslugi();
            Load_zakaz();
            Load_combobox();


            bunifuDataGridView3.Columns.Add("newColumnName", "ID");
            bunifuDataGridView3.Columns.Add("newColumnName1", "Название");
            bunifuDataGridView3.Columns.Add("newColumnName2", "Описание");
            bunifuDataGridView3.Columns.Add("newColumnName3", "Стоимость");

            bunifuDataGridView3.Columns[0].Visible = false;
            bunifuDataGridView3.Columns[2].Visible = false;

            bunifuDataGridView3.Columns[0].ValueType = typeof(int);
            bunifuDataGridView3.Columns[1].ValueType = typeof(string);
            bunifuDataGridView3.Columns[2].ValueType = typeof(string);
            bunifuDataGridView3.Columns[3].ValueType = typeof(int);
            bunifuDataGridView3.AllowUserToAddRows = false;
            bunifuDataGridView3.RowHeadersVisible = false;
        }

        //*****
        //*****  ===== Верхнее меню =====
        //*****
        //*****
        private void Main_users_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();//закрывать приложение при закрытии форму
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Application.Exit();//закрывать приложение при закрытии форму
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form1 fr = new Form1(); //возращаемся к форме авторизации
            fr.Show();
            this.Hide();//закрываем текущую форму
        }
        private void моиЗаявкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string name = ((ToolStripMenuItem)sender).Text;
            bunifuPages1.SetPage(name); //при нажатии на кнопку открывать соответствующую вкладку     
            Load_uslugi();
            Load_oboryd();
            switch (int.Parse(((ToolStripMenuItem)sender).Tag.ToString()))
            {
                case 1:
                    this.Text = "Окно пользователя - просмотр заявок";
                    title = this.Text;
                    texts = "Вы находитесь в окне просмотра Ваших заявок. Для получения иформации по заказу - нажмите на него. Для получения информаци о входящей в состав заказа услуги или оборудовании наведите на него курсор мыши";
                    break;
                case 2:
                    this.Text = "Окно пользователя - новая заявка";
                    title = this.Text;
                    texts = "Вы находитесь в окне добавления новой заявки. При добавлении новой заявки необходимо указать сведения об организации, название мероприятия, дату и время проведения мероприятия, контактный телефон, выбрать помещение и добавить необходимые услуги и оборудования";
                    break;
            }
        }

        //*****
        //*****  ===== END =====
        //*****
        //*****




        //*****
        //*****  ===== Загрузка информации =====
        //*****
        //*****
        private void Load_zakaz()
        {
            try  //перехват ошибок
            {
                // строка подключения к БД
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    MySqlCommand cmd = new MySqlCommand("SELECT Zakaz.ID, Zakaz.Corporation, USers.Name, Zakaz.name_dead, Zakaz.Time_start, Zakaz.Time_end,  Zakaz.Kolvo_person, cemetery.Name, Zakaz.Summa, Zakaz.Status, Zakaz.Nomer From zakaz INNER JOIN  USers ON Users.ID = Zakaz.Name  INNER JOIN  cemetery ON cemetery.ID = Zakaz.cemetery WHERE zakaz.Name = " + _id.ToString(), connection);
                    MySqlDataReader dr = cmd.ExecuteReader();
                    DataTable ds = new DataTable();
                    ds.Load(dr);                                                    //записываем данные с БД
                    bunifuDataGridView2.DataSource = ds;                        //выводим данные в форму
                    bunifuDataGridView2.RowHeadersVisible = false;                        //скрываем столбец с номерами строк
                    bunifuDataGridView2.Columns[0].Visible = false;               //убираем столбец с id
                    bunifuDataGridView2.Columns[1].HeaderText = "Название организации (при наличии)";  //название солбцов
                    bunifuDataGridView2.Columns[2].HeaderText = "Ф.И.О. заявителя";
                    bunifuDataGridView2.Columns[3].HeaderText = "Ф.И.О. покойного";
                    bunifuDataGridView2.Columns[4].HeaderText = "Время начала";
                    bunifuDataGridView2.Columns[5].HeaderText = "Время окончания";
                    bunifuDataGridView2.Columns[6].HeaderText = "Количество человек";
                    bunifuDataGridView2.Columns[7].HeaderText = "Кладбище";
                    bunifuDataGridView2.Columns[8].HeaderText = "Сумма к оплате";
                    bunifuDataGridView2.Columns[9].HeaderText = "Состояние";
                    bunifuDataGridView2.Columns[10].HeaderText = "Контактный телефон";
                    bunifuDataGridView2.AllowUserToAddRows = false;
                    bunifuDataGridView2.RowHeadersVisible = false;
                }

            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void Load_combobox()
        {
            try  //перехват ошибок
            {
                // строка подключения к БД
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    MySqlCommand cmd = new MySqlCommand("Select * From cemetery", connection);
                    MySqlDataReader dr = cmd.ExecuteReader();
                    pomes = new DataTable();
                    pomes.Load(dr);                                                    //записываем данные с БД
                    foreach (DataRow row in pomes.Rows) //заносим в список авторов
                    {
                        combo_pom.Add(new Data(int.Parse(row[0].ToString()), row[1].ToString()));
                    }
                    comboBox2.DataSource = combo_pom;
                    comboBox2.DisplayMember = "Name";
                    comboBox2.ValueMember = "id";
                }
                comboBox2.SelectedItem = comboBox2.Items[0];
                comboBox2_SelectedIndexChanged(new object(), new EventArgs());
            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void Load_uslugi()
        {
            try  //перехват ошибок
            {
                // строка подключения к БД
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    MySqlCommand cmd = new MySqlCommand("SELECT * FROM Uslugi", connection);
                    MySqlDataReader dr = cmd.ExecuteReader();
                    DataTable ds = new DataTable();
                    ds.Load(dr);//записываем данные с БД
                    bunifuDataGridView1.DataSource = ds;                        //выводим данные в форму
                    bunifuDataGridView1.RowHeadersVisible = false;                        //скрываем столбец с номерами строк
                    bunifuDataGridView1.Columns[0].Visible = false;               //убираем столбец с id
                    bunifuDataGridView1.Columns[1].HeaderText = "Название услуги";  //название солбцов
                    bunifuDataGridView1.Columns[3].HeaderText = "Стоимость";
                    bunifuDataGridView1.Columns[2].Visible = false;
                    bunifuDataGridView1.AllowUserToAddRows = false;
                    bunifuDataGridView1.RowHeadersVisible = false;
                }

                foreach (DataGridViewRow dr in bunifuDataGridView1.Rows)
                {
                    dr.Cells[1].ToolTipText = dr.Cells[2].Value.ToString();
                    dr.Cells[3].ToolTipText = dr.Cells[2].Value.ToString();
                }
                bunifuDataGridView1.ShowCellToolTips = true;

            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }
        private void Load_oboryd()
        {
            try  //перехват ошибок
            {
                // строка подключения к БД
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    MySqlCommand cmd = new MySqlCommand("SELECT * FROM Oboryd", connection);
                    MySqlDataReader dr = cmd.ExecuteReader();
                    DataTable ds = new DataTable();
                    ds.Load(dr);//записываем данные с БД
                    bunifuDataGridView5.DataSource = ds;                        //выводим данные в форму
                    bunifuDataGridView5.RowHeadersVisible = false;                        //скрываем столбец с номерами строк
                    bunifuDataGridView5.Columns[0].Visible = false;               //убираем столбец с id
                    bunifuDataGridView5.Columns[1].HeaderText = "Название оборудования";  //название солбцов
                    bunifuDataGridView5.Columns[3].HeaderText = "Стоимость";
                    bunifuDataGridView5.Columns[2].Visible = false;
                    bunifuDataGridView5.AllowUserToAddRows = false;
                    bunifuDataGridView5.RowHeadersVisible = false;
                }
                foreach (DataGridViewRow dr in bunifuDataGridView5.Rows)
                {
                    dr.Cells[1].ToolTipText = dr.Cells[2].Value.ToString();
                    dr.Cells[3].ToolTipText = dr.Cells[2].Value.ToString();
                }
            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void Load_dop_zakaz()
        {
            if (bunifuDataGridView2.Rows.Count > 0 && bunifuDataGridView2.SelectedRows.Count > 0)
                try  //перехват ошибок
                {
                    DataGridViewSelectedCellCollection DGVCell = bunifuDataGridView2.SelectedCells; //получаем номер выделенной строчки
                    var dgvc = DGVCell[1];//запоминаем данные из выделенной строки
                    var dgvr = dgvc.OwningRow;
                    using (var connection = new MySqlConnection(connectionString))
                    {
                        connection.Open();
                        MySqlCommand cmd = new MySqlCommand("SELECT Oboryd.Name, Oboryd.Price, Oboryd.Opisanie  FROM zakaz_oborud INNER JOIN Oboryd ON Oboryd.ID = zakaz_oborud.ID_oborud WHERE zakaz_oborud.ID_zakaz =" + dgvr.Cells[0].Value.ToString(), connection);
                        MySqlDataReader dr = cmd.ExecuteReader();
                        DataTable ds = new DataTable();
                        ds.Load(dr);                                                    //записываем данные с БД
                        bunifuDataGridView6.DataSource = ds;                        //выводим данные в форму
                        bunifuDataGridView6.RowHeadersVisible = false;                        //скрываем столбец с номерами строк
                        bunifuDataGridView6.Columns[0].HeaderText = "Название оборудования";  //название солбцов
                        bunifuDataGridView6.Columns[1].HeaderText = "Цена";
                        bunifuDataGridView6.Columns[2].Visible = false;
                        bunifuDataGridView6.AllowUserToAddRows = false;
                        bunifuDataGridView6.RowHeadersVisible = false;
                    }
                    foreach (DataGridViewRow dr in bunifuDataGridView6.Rows)
                    {
                        dr.Cells[0].ToolTipText = dr.Cells[2].Value.ToString();
                        dr.Cells[1].ToolTipText = dr.Cells[2].Value.ToString();
                    }
                    using (var connection = new MySqlConnection(connectionString))
                    {
                        connection.Open();
                        MySqlCommand cmd = new MySqlCommand("SELECT Uslugi.Name, Uslugi.Price, Uslugi.Opisanie FROM zakaz_uslug INNER JOIN Uslugi ON Uslugi.ID = zakaz_uslug.ID_uslug WHERE zakaz_uslug.ID_zakaz =" + dgvr.Cells[0].Value.ToString(), connection);
                        MySqlDataReader dr = cmd.ExecuteReader();
                        DataTable ds = new DataTable();
                        ds.Load(dr);                                                    //записываем данные с БД
                        bunifuDataGridView7.DataSource = ds;                        //выводим данные в форму
                        bunifuDataGridView7.RowHeadersVisible = false;                        //скрываем столбец с номерами строк
                        bunifuDataGridView7.Columns[0].HeaderText = "Название услуги";  //название солбцов
                        bunifuDataGridView7.Columns[1].HeaderText = "Цена";
                        bunifuDataGridView7.Columns[2].Visible = false;
                        bunifuDataGridView7.AllowUserToAddRows = false;
                        bunifuDataGridView7.RowHeadersVisible = false;
                    }
                    foreach (DataGridViewRow dr in bunifuDataGridView7.Rows)
                    {
                        dr.Cells[0].ToolTipText = dr.Cells[2].Value.ToString();
                        dr.Cells[1].ToolTipText = dr.Cells[2].Value.ToString();
                    }
                    DateTime date1 = DateTime.Parse(dgvr.Cells[4].Value.ToString());
                    DateTime date2 = DateTime.Parse(dgvr.Cells[5].Value.ToString());
                    var prodolshit = (date2 - date1).TotalHours;
                    foreach (DataGridViewRow dataRow in bunifuDataGridView6.Rows)
                    {
                        dataRow.Cells[1].Value = int.Parse(dataRow.Cells[1].Value.ToString()) * prodolshit;
                    }
                    foreach (DataGridViewRow dataRow in bunifuDataGridView7.Rows)
                    {
                        dataRow.Cells[1].Value = int.Parse(dataRow.Cells[1].Value.ToString()) * prodolshit;
                    }
                }
                catch (Exception ex) //возникает при ошибках
                {
                    MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                    MessageBox.Show(ex.StackTrace.ToString());
                }
        }

        private void Rashet_summy()
        {
            double summ = 0;
            int n = 0;
            if (int.TryParse(comboBox2.SelectedValue.ToString(), out n))
            {
                var row = pomes.Select("ID = " + n.ToString()).ToList();
                var prodolshit = Math.Round((dateTimePicker2.Value - dateTimePicker1.Value).TotalHours);
                summ = int.Parse(row[0][5].ToString());
                bunifuDataGridView3.Rows.Clear();
                foreach (DataRow dataRow in oborud)
                    bunifuDataGridView3.Rows.Add(dataRow.ItemArray);
                foreach (DataRow dataRow in uslugi)
                    bunifuDataGridView3.Rows.Add(dataRow.ItemArray);
                foreach (DataGridViewRow dr in bunifuDataGridView3.Rows)
                {
                    summ += int.Parse(dr.Cells[3].Value.ToString()) * prodolshit;
                    dr.Cells[3].Value = int.Parse(dr.Cells[3].Value.ToString()) * prodolshit;
                }
                textBox18.Text = summ.ToString();
            }

        }
        //*****
        //*****  ===== END =====
        //*****
        //*****


        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            if (bunifuDataGridView5.SelectedRows.Count > 0)
            {
                var datarow = ((DataRowView)bunifuDataGridView5.SelectedRows[0].DataBoundItem).Row;
                bunifuDataGridView3.Rows.Add(datarow.ItemArray);
                oborud.Add(datarow);
                Rashet_summy();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int n = 0;
            if (int.TryParse(comboBox2.SelectedValue.ToString(), out n))
                if (n > 0)
                {
                    var dgvc = pomes.Select("ID = " + n.ToString());
                    byte[] data = (byte[])dgvc[0][3];
                    MemoryStream ms = new MemoryStream(data);//считываем в потоке изображения и декодируем
                    Image returnImage = Image.FromStream(ms);
                    pictureBox1.BackgroundImage = returnImage;
                    textBox1.Text = dgvc[0][4].ToString();
                    label4.Text = "Тип места на кладище - " + dgvc[0][2].ToString() + "     Стоимсоть - " + dgvc[0][5].ToString() + " руб/м";
                    Rashet_summy();
                }
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            if (bunifuDataGridView1.SelectedRows.Count > 0)
            {
                var datarow = ((DataRowView)bunifuDataGridView1.SelectedRows[0].DataBoundItem).Row;
                bunifuDataGridView3.Rows.Add(datarow.ItemArray);
                uslugi.Add(datarow);
                Rashet_summy();
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > dateTimePicker2.Value)
            {
                MessageBox.Show("Введите корректную дату");
                dateTimePicker2.Value = dateTimePicker1.Value;
            }
            else
                Rashet_summy();
        }

        private void убратьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var point = bunifuDataGridView3.PointToClient(contextMenuStrip1.Bounds.Location);
                var info = bunifuDataGridView3.HitTest(point.X, point.Y);
                // Работаем с ячейкой
                var value = bunifuDataGridView3[info.ColumnIndex, info.RowIndex].OwningRow.Cells[1].Value.ToString();
                foreach (DataRow dataRow in oborud)
                    if (dataRow[1].Equals(value)) { oborud.Remove(dataRow); break; }
                foreach (DataRow dataRow in uslugi)
                    if (dataRow[1].Equals(value)) { uslugi.Remove(dataRow); break; }
                bunifuDataGridView3.Rows.RemoveAt(info.RowIndex);
                Rashet_summy();
            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            var point = bunifuDataGridView3.PointToClient(contextMenuStrip1.Bounds.Location);
            var info = bunifuDataGridView3.HitTest(point.X, point.Y);

            // Отменяем показ контекстного меню, если клик был не на ячейке
            if (info.RowIndex == -1 || info.ColumnIndex == -1)
            {
                e.Cancel = true;
            }
        }

        private void bunifuDataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            Load_dop_zakaz();
        }

        private void bunifuButton18_Click(object sender, EventArgs e)
        {
            try
            {
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();

                    var cmd = new MySqlCommand("INSERT INTO Zakaz (Corporation, Name, Name_dead, Time_start, Time_end, Kolvo_person, cemetery, Summa, Status, Nomer ) VALUES  (@CO, @name, @ev, @ts, @te, @kp, @pom, @sum, @st, @nom)", connection);
                    cmd.Parameters.Add(new MySqlParameter("@CO", textBox2.Text));
                    cmd.Parameters.Add(new MySqlParameter("@name", _id));
                    cmd.Parameters.Add(new MySqlParameter("@ev", textBox14.Text));
                    cmd.Parameters.Add(new MySqlParameter("@ts", dateTimePicker1.Value.ToString()));
                    cmd.Parameters.Add(new MySqlParameter("@te", dateTimePicker2.Value.ToString()));
                    cmd.Parameters.Add(new MySqlParameter("@kp", textBox17.Text));
                    cmd.Parameters.Add(new MySqlParameter("@pom", comboBox2.SelectedValue));
                    cmd.Parameters.Add(new MySqlParameter("@sum", textBox18.Text));
                    cmd.Parameters.Add(new MySqlParameter("@st", "Ожидание"));
                    cmd.Parameters.Add(new MySqlParameter("@nom", maskedTextBox1.Text));
                    cmd.ExecuteNonQuery();
                    Load_zakaz();
                }
                DataGridViewRow rows = bunifuDataGridView2.Rows[bunifuDataGridView2.Rows.Count - 1];
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    foreach (DataRow dr in oborud)
                    {
                        var cmd = new MySqlCommand("INSERT INTO zakaz_oborud (ID_zakaz, ID_oborud) VALUES (@zak, @obor)", connection);
                        cmd.Parameters.Add(new MySqlParameter("@zak", rows.Cells[0].Value));
                        cmd.Parameters.Add(new MySqlParameter("@obor", dr[0]));
                        cmd.ExecuteNonQuery();
                    }
                }
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    foreach (DataRow dr in uslugi)
                    {
                        var cmd = new MySqlCommand("INSERT INTO zakaz_uslug (ID_zakaz, ID_uslug) VALUES (@zak, @obor)", connection);
                        cmd.Parameters.Add(new MySqlParameter("@zak", rows.Cells[0].Value));
                        cmd.Parameters.Add(new MySqlParameter("@obor", dr[0]));
                        cmd.ExecuteNonQuery();
                    }
                }
                bunifuPages1.SetPage("Мои заявки"); //при нажатии на кнопку открывать соответствующую вкладку 
            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(texts, title);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult dialogResult = MessageBox.Show("Сохранить сведения о поданных заявках?", "Заявки", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    application = new Word.Application();
                    Object type = Type.Missing;
                    document = application.Documents.Add(ref type, ref _missingObj, ref _missingObj, ref _missingObj);
                    application.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                    Object start = 0;
                    Object end = 0;
                    Word.Range wordrange = document.Range(ref start, ref end);
                    wordrange.Text = "Сведения о поданных заявках \n по состоянию на " + DateTime.Now.ToLongDateString();
                    wordrange.Bold = 1;
                    wordrange.Font.Size = 14;
                    wordrange.Font.Name = "Times New Roman";
                    //Получаем ссылки на параграфы документа
                    var wordparagraphs = document.Paragraphs;
                    //Будем работать с первым параграфом
                    var wordparagraph = (Word.Paragraph)wordparagraphs[1];
                    wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordparagraph = (Word.Paragraph)wordparagraphs[2];
                    wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    start = wordrange.Text.Length;
                    end = wordrange.Text.Length;
                    wordrange = document.Range(ref start, ref end);
                    Object defaultTableBehavior =
                       Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                    //Добавляем таблицу и получаем объект wordtable 
                    Word.Table _table = document.Tables.Add(wordrange, 1, 11, ref defaultTableBehavior, ref autoFitBehavior);
                    var _currentRange = _table.Cell(1, 1).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "№ п/п";
                    _currentRange = _table.Cell(1, 2).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Организация";
                    _currentRange = _table.Cell(1, 3).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "ФИО заявителя";
                    _currentRange = _table.Cell(1, 4).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "ФИО покойного";

                    _currentRange = _table.Cell(1, 5).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Время начала";

                    _currentRange = _table.Cell(1, 6).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Время окончания";

                    _currentRange = _table.Cell(1, 7).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Количество персон";

                    _currentRange = _table.Cell(1, 8).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Кладбище";

                    _currentRange = _table.Cell(1, 9).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Сумма к оплате";

                    _currentRange = _table.Cell(1, 10).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Статус заказа";

                    _currentRange = _table.Cell(1, 11).Range;
                    _currentRange.Bold = 1;
                    _currentRange.Text = "Контактный телефон";

                    using (var connection = new MySqlConnection(connectionString))
                    {
                        connection.Open();
                        MySqlCommand cmd = new MySqlCommand("SELECT Zakaz.ID, Zakaz.Corporation, USers.Name, Zakaz.name_dead, Zakaz.Time_start, Zakaz.Time_end,  Zakaz.Kolvo_person, cemetery.Name, Zakaz.Summa, Zakaz.Status, Zakaz.Nomer From zakaz INNER JOIN  USers ON Users.ID = Zakaz.Name INNER JOIN  cemetery ON cemetery.ID = Zakaz.cemetery WHERE ZAkaz.Name = " + _id.ToString(), connection);
                        MySqlDataReader dr = cmd.ExecuteReader();
                        int i = 2;
                        while (dr.Read())
                        {
                            _table.Rows.Add(ref _missingObj);
                            _currentRange = _table.Cell(i, 1).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = (i - 1).ToString();
                            _currentRange = _table.Cell(i, 2).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[1].ToString();
                            _currentRange = _table.Cell(i, 3).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[2].ToString();
                            _currentRange = _table.Cell(i, 4).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[3].ToString();

                            _currentRange = _table.Cell(i, 5).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[4].ToString();

                            _currentRange = _table.Cell(i, 6).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[5].ToString();

                            _currentRange = _table.Cell(i, 7).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[6].ToString();

                            _currentRange = _table.Cell(i, 8).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[7].ToString();

                            _currentRange = _table.Cell(i, 9).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[8].ToString();

                            _currentRange = _table.Cell(i, 10).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[9].ToString();

                            _currentRange = _table.Cell(i, 11).Range;
                            _currentRange.Bold = 0;
                            _currentRange.Text = dr[10].ToString();

                            i++;
                        }
                        object begCell = _table.Cell(1, 1).Range.Start;
                        object endCell = _table.Cell(i--, 11).Range.End;
                        Word.Range wordcellrange = document.Range(ref begCell, ref endCell);
                        wordcellrange.Font.Size = 12;
                        wordcellrange.Font.Name = "Times New Roman";
                        wordparagraph = (Word.Paragraph)wordparagraphs[3];
                        wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }


                var sfd = new SaveFileDialog() { Filter = "Word Documents (.docx)|*.docx|Word Template (.dotx)|*.dotx" };
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // получаем выбранный файл
                    Object pathToSaveObj = saveFileDialog1.FileName;
                    document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj);
                    application.Visible = true;
                }


            }
            catch (Exception ex) //возникает при ошибках
            {
                document.Close(ref falseObj, ref _missingObj, ref _missingObj);
                application.Quit(ref _missingObj, ref _missingObj, ref _missingObj);
                document = null;
                application = null;
            }
            finally
            {
                document.Close(ref falseObj, ref _missingObj, ref _missingObj);
                application.Quit(ref _missingObj, ref _missingObj, ref _missingObj);
                document = null;
                application = null;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void вExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Сохранить сведения о поданных заявках?", "Заявки", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    //Объявляем приложение
                    app = new Excel.Application
                    {
                        //Отобразить Excel
                        //  Visible = true,
                        //Количество листов в рабочей книге
                        SheetsInNewWorkbook = 1
                    };
                    //Добавить рабочую книгу
                    workBook = app.Workbooks.Add(Type.Missing);
                    //Отключить отображение окон с сообщениями
                    app.DisplayAlerts = false;
                    //Получаем первый лист документа (счет начинается с 1)
                    sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);
                    //Название листа (вкладки снизу)
                    sheet.Name = "Заявки";

                    sheet.Range["A1"].Value = "Сведения о поданных заявках \n по состоянию на " + DateTime.Now.ToLongDateString();
                    Excel.Range range2 = sheet.get_Range("A1", "K1");
                    range2.Merge(Type.Missing);
                    range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range2.Cells.Font.Name = "Times New Roman";
                    //Размер шрифта для диапазона
                    range2.Cells.Font.Size = 14;
                    //Жирный текст
                    range2.Font.Bold = true;


                    sheet.Range["A2"].Value = "№ п/п";
                    sheet.Range["B2"].Value = "Организация";
                    sheet.Range["C2"].Value = "ФИО заявителя";
                    sheet.Range["D2"].Value = "ФИО покойного";
                    sheet.Range["E2"].Value = "Время начала";
                    sheet.Range["F2"].Value = "Время окончания";
                    sheet.Range["G2"].Value = "Количество персон";
                    sheet.Range["H2"].Value = "Кладбище";
                    sheet.Range["I2"].Value = "Сумма к оплате";
                    sheet.Range["J2"].Value = "Статус заказа";
                    sheet.Range["K2"].Value = "Контактный телефон";
                    range2 = sheet.get_Range("A2", "K2");
                    range2.Cells.Font.Name = "Times New Roman";
                    //Размер шрифта для диапазона
                    range2.Cells.Font.Size = 12;
                    //Жирный текст
                    range2.Font.Bold = true;
                    int i = 3;
                    using (var connection = new MySqlConnection(connectionString))
                    {
                        connection.Open();
                        MySqlCommand cmd = new MySqlCommand("SELECT Zakaz.ID, Zakaz.Corporation, USers.Name, Zakaz.name_dead, Zakaz.Time_start, Zakaz.Time_end,  Zakaz.Kolvo_person, cemetery.Name, Zakaz.Summa, Zakaz.Status, Zakaz.Nomer From zakaz INNER JOIN  USers ON Users.ID = Zakaz.Name INNER JOIN  cemetery ON cemetery.ID = Zakaz.cemetery WHERE ZAkaz.Name = " + _id.ToString(), connection);
                        MySqlDataReader dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            sheet.Range["A" + (i).ToString()].Value = (i - 2).ToString();
                            sheet.Range["B" + (i).ToString()].Value = dr[1].ToString();
                            sheet.Range["C" + (i).ToString()].Value = dr[2].ToString();
                            sheet.Range["D" + (i).ToString()].Value = dr[3].ToString();
                            sheet.Range["E" + (i).ToString()].Value = dr[4].ToString();
                            sheet.Range["F" + (i).ToString()].Value = dr[5].ToString();
                            sheet.Range["G" + (i).ToString()].Value = dr[6].ToString();
                            sheet.Range["H" + (i).ToString()].Value = dr[7].ToString();
                            sheet.Range["I" + (i).ToString()].Value = dr[8].ToString();
                            sheet.Range["J" + (i).ToString()].Value = dr[9].ToString();
                            sheet.Range["K" + (i).ToString()].Value = dr[10].ToString();
                            i++;
                        }
                    }

                    Excel.Range range = sheet.get_Range("A3", "K" + (i - 1).ToString());
                    range.Cells.Font.Name = "Times New Roman";
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //Размер шрифта для диапазона
                    range.Cells.Font.Size = 12;
                    range.EntireColumn.AutoFit();
                    range.Borders.Color = ColorTranslator.ToOle(Color.Black);

                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Excel files(*.xls)|*.xls";    //формат выходных файлов
                    //Сохраняем файл
                    if (save.ShowDialog() == DialogResult.Cancel)
                        return;
                    // получаем выбранный файл
                    string filename = save.FileName;
                    app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    sheet = null;
                    app = null;
                    workBook.Close();
                }
                catch (Exception ex) //возникает при ошибках
                {
                    MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                    MessageBox.Show(ex.StackTrace.ToString());
                }
            }
        }
            
    }
}


