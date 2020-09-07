using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using A.E.SCryptLibrary;

namespace CarParts
{
    public partial class Authentification : Form
    {
        public string SecretPass = "3310";
        public OdbcConnection connection = new OdbcConnection(Properties.Settings.Default.vvgcarpartsConnectionString);
        public static string userLogin;


        public Authentification()
        {
            InitializeComponent();
        }

        private void Authentification_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void RegisterPanelButton_Click(object sender, EventArgs e)
        {
            RegistrationPanel.Location = new Point(574, 281);
            RegistrationPanel.Visible = true;
            AuthentificationPanel.Visible = false;
        }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            RegistrationPanel.Visible = false;
            AuthentificationPanel.Visible = true;
            AuthentificationPanel.Location = new Point(574, 353);
        }

        private void RegistrationButton_Click(object sender, EventArgs e)
        {
            IPStatus status = IPStatus.Unknown;
            try
            {
                status = new Ping().Send(@"yandex.ru").Status;
            }
            catch { }

            if (status != IPStatus.Success)
            {
                ThrowICFError("Отсутствует интернет соединение. Проверьте подключение к сети.");
                return;
            }
            if (LoginRTextBox.Text == "")
            {
                ThrowError("Поле логина для регистрации пустое, повторите ввод.");
                return;
            }
            if (PassRTextBox.Text == "")
            {
                ThrowError("Поле пароля для регистрации пустое, повторите ввод.");
                return;
            }
            if (SecretRTextBox.Text == "")
            {
                ThrowError("Поле секретного пароля для регистрации пустое, повторите ввод.");
                return;
            }
            if (AgainPassRTextBox.Text == "")
            {
                ThrowError("Поле повтора пароля для регистрации пустое, повторите ввод.");
                return;
            }
            if (AgainPassRTextBox.Text != PassRTextBox.Text)
            {
                ThrowError("Пароли не совпадают, проверьте правильность введённых данных.");
                return;
            }
            if (SecretRTextBox.Text != SecretPass)
            {
                ThrowError("Не правильный секретный пароль, обратитесь за помощью к администратору.");
                return;
            }
            if (PassRTextBox.Text.Length < 5)
            {
                ThrowError("Пароль слишком короткий. Он должен содержать не менее 5 символов.");
                return;
            }
            if (!CheckSymbols(PassRTextBox.Text))
            {
                ThrowError("Пароль содержит недопустимые символы. Повторите ввод.");
                return;
            }

            BagCrypt bc = new BagCrypt();

            connection.Open();
            var command = new OdbcCommand($"insert into users values(default, '{LoginRTextBox.Text}', '{bc.CryptBag(PassRTextBox.Text)}')", connection);
            command.ExecuteNonQuery();
            connection.Close();
            ShowSuccess("Вы успешно зарегистрировались в системе. Перейдите к панели авторизации.");
            ClearRegFields();
        }


        public bool CheckSymbols(string text)
        {
            bool flag = false;
            var arr = text.ToCharArray();
            for (int i = 0; i < arr.Length; i++)
            {
                if (!char.IsLetterOrDigit(arr[i]))
                {
                    flag = false;
                }
                else
                {
                    flag = true;
                }
            }
            return flag;
        }

        public void ClearRegFields()
        {
            LoginRTextBox.Text = PassRTextBox.Text = AgainPassRTextBox.Text = SecretRTextBox.Text = "";
        }

        public void ThrowError(string message)
        {
            ErrorPanel.Visible = true;
            ErrorPanel.Location = new Point(636, 398);
            ErrorRichBox.Text = message;
        }

        public void ThrowICFError(string message)
        {
            InternetConnectionFailedPanel.Visible = true;
            ICFRichBox.Text = message;
        }

        public void ShowSuccess(string message)
        {
            SuccessPanel.Visible = true;
            SuccessPanel.Location = new Point(636, 398);
            SuccessRichBox.Text = message;
        }

        private void OkErrorButton_Click(object sender, EventArgs e)
        {
            ErrorPanel.Visible = false;
        }

        private void OkSuccessButton_Click(object sender, EventArgs e)
        {
            SuccessPanel.Visible = false;
        }

        private void EnterProgramButton_Click(object sender, EventArgs e)
        {
            IPStatus status = IPStatus.Unknown;
            try
            {
                status = new Ping().Send(@"yandex.ru").Status;
            }
            catch { }

            if (status != IPStatus.Success)
            {
                ThrowICFError("Отсутствует интернет соединение. Проверьте подключение к сети.");
                return;
            }

            BagCrypt bc = new BagCrypt();

            bool CheckUser(string login, string pass)
            {
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                var check = new OdbcCommand($"select 1 from users where userLogin = '{login}' and userPassword = '{bc.CryptBag(pass)}';", connection);
                return 1 == Convert.ToInt32(check.ExecuteScalar());
            }

            if (LoginATextBox.Text == "")
            {
                ThrowError("Поле логина для входа пустое, повторите ввод.");
                return;
            }
            if (PassATextBox.Text == "")
            {
                ThrowError("Поле пароля для входа пустое, повторите ввод.");
                return;
            }
            if (!CheckUser(LoginATextBox.Text, PassATextBox.Text))
            {
                ThrowError($"Не правильный ввод данных, проверьте введённый логин и пароль.");
                connection.Close();
                return;
            }
            if (CheckUser(LoginATextBox.Text, PassATextBox.Text))
            {
                Main mn = new Main();
                mn.Owner = this;
                mn.Show();
            }
        }

        private void LoginATextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void ExitProgramButton_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void ICFOkButton_Click(object sender, EventArgs e)
        {
            InternetConnectionFailedPanel.Visible = false;
        }
    }
}
