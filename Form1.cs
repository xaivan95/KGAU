using MySql.Data.MySqlClient;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace KGAU
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            bunifuTextBox2.UseSystemPasswordChar = true; //скрываем вводимые данные в поле пароль
            bunifuTextBox2.PasswordChar = '*';//символ вместо вводимых символов
        }
        public string name;
        public int role;
        public int id;
        private bool isMousePress = false; //переменная хранит нажатия кнопки мыши
        private Point _clickPoint; //точка мыши
        private Point _formStartPoint;//точка начала
        string connectionString = "server=localhost;user=root;database=ritualdb;password=2244;";
        //кнопка Войти
        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            try  //перехват ошибок
            {

                MySqlConnection connection = new MySqlConnection(String.Format(connectionString, bunifuTextBox1.Text, bunifuTextBox2.Text));
                connection.Open();
                string query = "SELECT * FROM Users WHERE Login = '" + bunifuTextBox1.Text + "'";
                 MySqlCommand com = new MySqlCommand(query, connection);
                bool flag = false;
                using (MySqlDataReader dr = com.ExecuteReader())                  
                {
                    while (dr.Read())
                    {
                        if (dr[3].ToString().Equals(bunifuTextBox2.Text))
                        {
                            name = dr[1].ToString();//запоминаем имя
                            role = int.Parse(dr[4].ToString());//роль
                            id = int.Parse(dr[0].ToString());//id
                            flag = true;
                        }
                    }
                }
                connection.Close();
               
                if (flag)//если логин и пароль уникальны и совпали
                {
                    switch (role) //исходя из роли пользователя
                    {//                 загружаем определенную форму
                        case 1:
                            Main_admin fr = new Main_admin();
                            fr.Show();
                            this.Hide();
                            break;
                        case 2:
                            Main_users frk = new Main_users(id);
                            frk.Show();
                            this.Hide();
                            break;

                    }
                }



            }
            catch (Exception ex) //возникает при ошибках
            {
                MessageBox.Show("Ошибка подключения: " + ex.Message.ToString()); //выводим полученную ошибку
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }

        private void bunifuGradientPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMousePress)
            {
                var cursorOffsetPoint = new Point( //считаем смещение курсора от старта
                    Cursor.Position.X - _clickPoint.X,
                    Cursor.Position.Y - _clickPoint.Y);

                Location = new Point( //смещаем форму от начальной позиции в соответствии со смещением курсора
                    _formStartPoint.X + cursorOffsetPoint.X,
                    _formStartPoint.Y + cursorOffsetPoint.Y);
            }
        }

        private void bunifuGradientPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            
                isMousePress = true; //запомнили что кнопка нажата
                _clickPoint = Cursor.Position; //запомнили позиции мышки
                _formStartPoint = Location;
        }

        private void bunifuGradientPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            isMousePress = false;//запоминаем что клавиша мыши отпущена
            _clickPoint = Point.Empty;
        }

        private void bunifuTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) //если нажата кнопка ENTER
            {
                bunifuButton1_Click(sender, e); //тоже самое что нажать кнопку ВОЙТИ
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Application.Exit();//Выход  
        }
    }
}
