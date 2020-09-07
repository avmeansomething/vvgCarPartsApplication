using System;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Protobuf.Collections;
using System.Data.Odbc;
using System.Net.Http;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.VisualBasic;
using System.Net.NetworkInformation;
using System.Net;

namespace CarParts
{
    public partial class Main : Form
    {
        public OdbcConnection connection = new OdbcConnection(Properties.Settings.Default.vvgcarpartsConnectionString);
        public Main()
        {
            connection.Open();
            InitializeComponent();
        }
        public bool InternetConnection = true;
        public virtual void FillByAutoName(ComboBox box)
        {
            var command = new OdbcDataAdapter($"select distinct auto_name from autos;", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_name";
        }
        public virtual void FillByAutoName(ComboBox box, string partName)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}';", connection);
            if (connection.State == ConnectionState.Open)
            {
                int id = Convert.ToInt32(getId.ExecuteScalar());

                var command = new OdbcDataAdapter($"select distinct auto_name from autos where id = '{id}';", connection);
                var columns = new OdbcCommandBuilder(command);
                DataSet ds = new DataSet();
                command.Fill(ds, "autos");
                box.DataSource = ds.Tables[0];
                box.DisplayMember = "auto_name";
            }
        }
        public virtual void FillByAutoNameByID(ComboBox box, string car_id)
        {
            var command = new OdbcDataAdapter($"select distinct auto_name from autos where id = '{car_id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_name";
        }
        public void FillByAutoNameDet(ComboBox box, string partName)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}';", connection);
            if (connection.State == ConnectionState.Open)
            {
                int id = Convert.ToInt32(getId.ExecuteScalar());

                var command = new OdbcDataAdapter($"select distinct auto_name from autos;", connection);
                var columns = new OdbcCommandBuilder(command);
                DataSet ds = new DataSet();
                command.Fill(ds, "autos");
                box.DataSource = ds.Tables[0];
                box.DisplayMember = "auto_name";
            }
        }
        public virtual void FillByAutoModel(ComboBox box, string autoName)
        {
            var command = new OdbcDataAdapter($"select distinct auto_model from autos where auto_name = '{autoName}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_model";
        }
        public virtual void FillByAutoModelForParts(ComboBox box, string partName)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}';", connection);
            int id = Convert.ToInt32(getId.ExecuteScalar());
            var command = new OdbcDataAdapter($"select distinct auto_model from autos where id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_model";
        }
        public virtual void FillByAutoFuel(ComboBox box, string autoName, string autoModel)
        {
            var command = new OdbcDataAdapter($"select distinct auto_fuelvolume from autos where auto_name = '{autoName}' and auto_model = '{autoModel}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_fuelvolume";
        }
        public virtual void FillByAutoYear(ComboBox box, string autoName, string autoModel, string autoFuel)
        {
            var command = new OdbcDataAdapter($"select distinct auto_year from autos where auto_name = '{autoName}' and auto_model = '{autoModel}' and auto_fuelvolume = '{autoFuel}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_year";
        }
        public virtual void FillByAutoYear(ComboBox box, string partName)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}';", connection);
            int id = Convert.ToInt32(getId.ExecuteScalar());
            var command = new OdbcDataAdapter($"select distinct auto_year from autos where id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_year";
        }
        public virtual void FillByAutoFuel(ComboBox box, string partName)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}' and part_code = '';", connection);
            id = getId.ExecuteScalar().ToString();
            var command = new OdbcDataAdapter($"select distinct auto_fuelvolume from autos where id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_fuelvolume";
        }
        public void FillByAutoFuelCode(ComboBox box, string partName, string partCode)
        {
            var getId = new OdbcCommand($"select car_id from parts where part_name = '{partName}' and part_code = '{partCode}';", connection);
            id = getId.ExecuteScalar().ToString();
            var command = new OdbcDataAdapter($"select distinct auto_fuelvolume from autos where id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "autos");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "auto_fuelvolume";
        }
        public void FillByPartName(ComboBox box)
        {
            var command = new OdbcDataAdapter($"select distinct part_name from parts;", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "parts");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "part_name";
        }
        public void FillByPartDescription(ComboBox box, string partName)
        {
            var command = new OdbcDataAdapter($"select distinct part_description from parts where part_name = '{partName}' and car_id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "parts");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "part_description";
        }
        public void FillByPartCode(ComboBox box, string partName)
        {
            var command = new OdbcDataAdapter($"select distinct part_code from parts where part_name = '{partName}';", connection);
            var columns = new OdbcCommandBuilder(command);
            DataSet ds = new DataSet();
            command.Fill(ds, "parts");
            box.DataSource = ds.Tables[0];
            box.DisplayMember = "part_code";
        }
        public void FillByCategoryName(ComboBox box)
        {
            var command = new OdbcDataAdapter($"select distinct category_name from categories;", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "categories");

            box.DataSource = ds.Tables[0];
            box.DisplayMember = "category_name";
        }
        public string GetCategoryName(string id)
        {
            var command = new OdbcCommand($"select category_name from categories where id = '{id}';", connection);
            return command.ExecuteScalar().ToString();
        }
        public void FillByCategoryName(ComboBox box, string id)
        {
            var command = new OdbcDataAdapter($"select distinct category_name from categories where id = '{id}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "categories");

            box.DataSource = ds.Tables[0];
            box.DisplayMember = "category_name";
        }
        public virtual string ReturnAutoId(string autoName, string autoModel, string autoFuelType, string autoYear)
        {
            var command = new OdbcCommand($"select id from autos where auto_name = '{autoName}' and auto_model = '{autoModel}' and auto_fuelvolume = '{autoFuelType}' and auto_year = '{autoYear}';", connection);
            return command.ExecuteScalar().ToString();
        }
        public string id = string.Empty;
        public void ThrowError(string message)
        {
            ErrorPanel.Visible = true;
            ErrorPanel.Location = new System.Drawing.Point(636, 398);
            ErrorRichBox.Text = message;
            ErrorPanel.BringToFront();
        }
        public void AddNewRowStrip(string message)
        {
            NewRowPanel.Visible = true;
            NewRowPanel.Location = new System.Drawing.Point(636, 398);
            NewAutoRichBox.Text = message;
            NewRowPanel.BringToFront();
        }
        public void ShowSuccess(string message)
        {
            SuccessPanel.Visible = true;
            SuccessPanel.Location = new System.Drawing.Point(636, 398);
            SuccessRichBox.Text = message;
            SuccessPanel.BringToFront();
        }
        public void ConfirmMove(string message)
        {
            ConfirmPanel.Visible = true;
            ConfirmPanel.Location = new System.Drawing.Point(636, 398);
            ConfirmRichBox.Text = message;
            ConfirmPanel.BringToFront();
        }
        public bool CheckInternetConnection()
        {
            try
            {
                HttpWebRequest Request = (HttpWebRequest)WebRequest.Create("https://vk.com/");
                HttpWebResponse Response = (HttpWebResponse)Request.GetResponse();
                InternetConnection = true;
            }
            catch (WebException ex)
            {
                ThrowICFError("Отсутствует интернет соединение. Проверьте подключение к сети.");
                InternetConnection = false;
            }
            return InternetConnection;
        }
        private void Main_Load(object sender, EventArgs e)
        {
            this.usersTableAdapter.Fill(this.vvgcarpartsDataSet.users);
            this.reviewsTableAdapter.Fill(this.vvgcarpartsDataSet.reviews);
            this.partsTableAdapter.Fill(this.vvgcarpartsDataSet.parts);
            this.part_photosTableAdapter.Fill(this.vvgcarpartsDataSet.part_photos);
            this.categoriesTableAdapter.Fill(this.vvgcarpartsDataSet.categories);
            this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
            FillByPartName(PartNameSellsComboBox);
            Authentification auth = this.Owner as Authentification;
            if (auth != null)
            {
                CurrentUserLabel.Text += auth.LoginATextBox.Text;
            }
            FillByAutoName(AutoNameComboBox);

            //RandomAutoDiagramm();
            AutoDataGrid.Focus();
        }
        private void ExitLabel_MouseEnter(object sender, EventArgs e)
        {
            ExitLabel.ForeColor = Color.FromArgb(82, 171, 133);
        }
        //public void AmountOfCarsDiagramm()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < AutoDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(AutoDataGrid.Rows[i].Cells[1].Value.ToString()))
        //            yValues.Add(AutoDataGrid.Rows[i].Cells[1].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(id) from autos where auto_name = '{yValues[i]}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По марке авто", xValues, yValues);
        //}
        //public void PartsCategoryCount()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < PartsDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(PartsDataGrid.Rows[i].Cells[3].Value.ToString()))
        //            yValues.Add(PartsDataGrid.Rows[i].Cells[3].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(ID) from partsview where CName = '{yValues[i]}';", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По категориям запчастей", xValues, yValues);
        //}
        //public void AmountOfYear()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < AutoDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(AutoDataGrid.Rows[i].Cells[6].Value.ToString()))
        //            yValues.Add(AutoDataGrid.Rows[i].Cells[6].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(id) from autos where auto_year = '{yValues[i]}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По годам выпуска", xValues, yValues);
        //}
        //public void AmountOfTransmisson()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < AutoDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(AutoDataGrid.Rows[i].Cells[5].Value.ToString()))
        //            yValues.Add(AutoDataGrid.Rows[i].Cells[5].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(id) from autos where auto_transmissiotype = '{yValues[i]}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По типу коробки передач", xValues, yValues);
        //}
        //public void AmountOfDriveType()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < AutoDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(AutoDataGrid.Rows[i].Cells[7].Value.ToString()))
        //            yValues.Add(AutoDataGrid.Rows[i].Cells[7].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(id) from autos where auto_drivetype = '{yValues[i]}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По типу ведущего привода", xValues, yValues);
        //}
        //public void MonthSells()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < PartPhotosDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(ReturnStringMonth(ReturnMonth(PartPhotosDataGrid.Rows[i].Cells[5].Value.ToString()))))
        //            yValues.Add(ReturnStringMonth(ReturnMonth(PartPhotosDataGrid.Rows[i].Cells[5].Value.ToString())));
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(sells_id) from sells where month(sells_date) = '{ReturnNumberMonth(yValues[i])}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По месяцу продажи", xValues, yValues);
        //}
        //public void AmountOfBody()
        //{
        //    var xValues = new List<int>();
        //    var yValues = new List<string>();
        //    for (int i = 0; i < AutoDataGrid.RowCount; i++)
        //    {
        //        if (!yValues.Contains(AutoDataGrid.Rows[i].Cells[3].Value.ToString()))
        //            yValues.Add(AutoDataGrid.Rows[i].Cells[3].Value.ToString());
        //    }
        //    for (int i = 0; i < yValues.Count; i++)
        //    {
        //        var command = new OdbcCommand($"select count(id) from autos where auto_bodytype = '{yValues[i]}'", connection);
        //        xValues.Add(Convert.ToInt32(command.ExecuteScalar()));
        //    }
        //    DrawPie("По типу кузова", xValues, yValues);
        //}
        //private void DrawPie(string name, List<int> xValues, List<string> yValues)
        //{
        //    MainDiagramm.Series.Clear();
        //    MainDiagramm.BackColor = Color.FromArgb(30, 119, 176);

        //    MainDiagramm.ChartAreas[0].BackColor = Color.FromArgb(37, 51, 64);

        //    MainDiagramm.Titles.Clear();
        //    MainDiagramm.Titles.Add(name);
        //    MainDiagramm.Titles[0].ForeColor = Color.White;
        //    MainDiagramm.Titles[0].Font = new System.Drawing.Font("Century Gothic", 16, FontStyle.Bold);

        //    MainDiagramm.Series.Add(new System.Windows.Forms.DataVisualization.Charting.Series("ColumnSeries")
        //    {
        //        ChartType = SeriesChartType.Pie
        //    });

        //    MainDiagramm.Series["ColumnSeries"].Points.DataBindXY(yValues, xValues);

        //    MainDiagramm.Series[0].Font = new System.Drawing.Font("Century Gothic", 10, FontStyle.Bold);
        //    MainDiagramm.Series[0].LabelForeColor = Color.White;
        //    MainDiagramm.Series[0].IsValueShownAsLabel = true;
        //    MainDiagramm.ChartAreas[0].Area3DStyle.Enable3D = true;
        //}
        private void ExitLabel_Click(object sender, EventArgs e)
        {
            Authentification auth = this.Owner as Authentification;
            if (auth != null)
            {
                auth.LoginATextBox.Text = "";
                auth.PassATextBox.Text = "";
            }
            this.Close();
        }
        private void ExitLabel_MouseLeave(object sender, EventArgs e)
        {
            ExitLabel.ForeColor = Color.White;
        }
        private void AutoTableButton_Click(object sender, EventArgs e)
        {
            HidingPanel.BringToFront();
            //RandomAutoDiagramm();
            AutoDataGrid.Visible = true;
            PartsDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
            AutoDataGrid.Focus();
            HideAllAdvPanels();
        }
        //private void RandomAutoDiagramm()
        //{
        //    var rnd = new Random();
        //    var num = rnd.Next(1, 6);
        //    switch (num)
        //    {
        //        case 1:
        //            AmountOfCarsDiagramm();
        //            break;
        //        case 2:
        //            AmountOfTransmisson();
        //            break;
        //        case 3:
        //            AmountOfYear();
        //            break;
        //        case 4:
        //            AmountOfBody();
        //            break;
        //        case 5:
        //            AmountOfDriveType();
        //            break;
        //    }
        //}
        private void PartsTableButton_Click(object sender, EventArgs e)
        {
            PartsDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            this.partsTableAdapter.Fill(this.vvgcarpartsDataSet.parts);
            PartsDataGrid.Focus();
            ConfirmPanel.Visible = false;
            HideAllAdvPanels();
            //PartsCategoryCount();
        }
        private void UsersTableButton_Click(object sender, EventArgs e)
        {
            UsersDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            PartsDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            UsersDataGrid.Focus();
            ConfirmPanel.Visible = false;
            HideAllAdvPanels();
            this.usersTableAdapter.Fill(this.vvgcarpartsDataSet.users);
        }
        private void CategoriesTableButton_Click(object sender, EventArgs e)
        {
            CategoriesDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            PartsDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            this.categoriesTableAdapter.Fill(this.vvgcarpartsDataSet.categories);
            CategoriesDataGrid.Focus();
            ConfirmPanel.Visible = false;
            HideAllAdvPanels();
        }
        private void PartsDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            if (PartsDataGrid.Visible && PartsDataGrid.CurrentRow != null)
            {
                MainPictureBox.Image = Properties.Resources.load1;

                CurrentItemLabel.Text = PartsDataGrid.CurrentRow.Cells[1].Value.ToString() + " " + PartsDataGrid.CurrentRow.Cells[4].Value.ToString() + " " + PartsDataGrid.CurrentRow.Cells[5].Value.ToString();
            }
        }
        private void AutoDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible && AutoDataGrid.CurrentRow != null && AutoDataGrid.DataSource != null)
            {
                MainPictureBox.Image = Properties.Resources.load1;
                MainPictureBox.ImageLocation = AutoDataGrid.CurrentRow.Cells[9].Value.ToString();
                CurrentItemLabel.Text = AutoDataGrid.CurrentRow.Cells[1].Value.ToString() + " " + AutoDataGrid.CurrentRow.Cells[2].Value.ToString() + " " + AutoDataGrid.CurrentRow.Cells[4].Value.ToString() + " " + AutoDataGrid.CurrentRow.Cells[5].Value.ToString() + " " + AutoDataGrid.CurrentRow.Cells[6].Value.ToString();
            }
        }
        private void AddVehicleButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(AutoAddPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            var insert = new OdbcCommand($"insert into autos values (default, '{AutoNameAddTextBox.Text}', '{AutoModelAddTextBox.Text}', '{AutoBodyTypeAddTextBox.Text}', '{AutoFuelTypeAddTextBox.Text}', '{AutoFuelVolumeAddTextBox.Text}', '{TransmissionAddTextBox.Text}', {AutoYearAddTextBox.Text}, '{DriveTypeAddTextBox.Text}', '{AutoPhotoAddTextBox.Text}');", connection);
            insert.ExecuteNonQuery();

            ShowSuccess("Данные в таблицу автомобилей успешно добавлены.");
            this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
        }
        public void CancelAddInfo(Panel panel)
        {
            var controlCollection = panel.Controls;
            for (int i = 0; i < controlCollection.Count; i++)
            {
                if (controlCollection[i].GetType() == typeof(System.Windows.Forms.TextBox))
                {
                    controlCollection[i].Text = "";
                }
            }
        }
        public void OffButtons()
        {
            AddInfoButton.Enabled = false;
            DeleteInfoButton.Enabled = false;
            FindInfoButton.Enabled = false;
            EditRowsButton.Enabled = false;
        }
        public void OnButtons()
        {
            AddInfoButton.Enabled = true;
            DeleteInfoButton.Enabled = true;
            FindInfoButton.Enabled = true;
            EditRowsButton.Enabled = true;
        }
        private void CancelAddingAutoButton_Click(object sender, EventArgs e)
        {
            CancelAddInfo(AutoAddPanel);
            AutoAddPanel.Visible = false;
            OnButtons();
        }
        private void AddInfoButton_Click(object sender, EventArgs e)
        {
            CancelAddInfo(AutoAddPanel);
            CancelAddInfo(PartsAddPanel);

            if (AutoDataGrid.Visible)
            {
                AutoAddPanel.Visible = true;
                AutoAddPanel.Location = new System.Drawing.Point(546, 106);
                OffButtons();
            }
            if (PartsDataGrid.Visible)
            {
                PartsAddPanel.Visible = true;
                PartsAddPanel.Location = new System.Drawing.Point(546, 106);
                OffButtons();
            }
            if (PartPhotosDataGrid.Visible)
            {
                FillByPartName(PartNameSellsComboBox);
                PartPhotoAddPanel.Visible = true;
                PartPhotoAddPanel.Location = new System.Drawing.Point(575, 175);
                OffButtons();
            }
            if (CategoriesDataGrid.Visible)
            {
                CategoryAddPanel.Visible = true;
                CategoryAddPanel.Location = new System.Drawing.Point(575, 395);
                OffButtons();
            }
        }
        private void OkErrorButton_Click(object sender, EventArgs e)
        {
            ErrorPanel.Visible = false;
        }
        private void OkSuccessButton_Click(object sender, EventArgs e)
        {
            SuccessPanel.Visible = false;
        }
        private void UpdateCellsGrid_Click(object sender, EventArgs e)
        {
            this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
            this.categoriesTableAdapter.Fill(this.vvgcarpartsDataSet.categories);
            this.usersTableAdapter.Fill(this.vvgcarpartsDataSet.users);
            this.part_photosTableAdapter.Fill(this.vvgcarpartsDataSet.part_photos);
            this.reviewsTableAdapter.Fill(this.vvgcarpartsDataSet.reviews);
        }
        private void AutoDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void CancelAddingPartsButton_Click(object sender, EventArgs e)
        {
            CancelAddInfo(PartsAddPanel);
            PartsAddPanel.Visible = false;
            OnButtons();

            AutoModelComboBox.DataSource = null;
            AutoModelComboBox.Text = string.Empty;
            AutoYearComboBox.DataSource = null;
            AutoYearComboBox.Text = string.Empty;
            AutoFuelTypeComboBox.DataSource = null;
            AutoFuelTypeComboBox.Text = string.Empty;

            AutoNameComboBox.SelectedIndex = 0;
        }
        public bool CheckFields(Panel panel)
        {
            var controlCollection = panel.Controls;
            bool flag = true;
            for (int i = 0; i < controlCollection.Count; i++)
            {
                if (controlCollection[i].GetType() == typeof(System.Windows.Forms.TextBox) || controlCollection[i].GetType() == typeof(ComboBox))
                {
                    if (controlCollection[i].Text == string.Empty)
                        flag = false;
                }
            }
            return flag;
        }
        public int ReturnCategoryId(string categoryName)
        {
            var command = new OdbcCommand($"select id from categories where category_name = '{categoryName}' ;", connection);
            return Convert.ToInt32(command.ExecuteScalar());
        }
        public string ReturnPartId(string partName, string carId, string partCode)
        {
            var command = new OdbcCommand($"select part_id from parts where part_name = '{partName}' and car_id = '{carId}' and part_code = '{partCode}';", connection);
            return command.ExecuteScalar().ToString();
        }
        public virtual string ReturnPartId(string partName, string partCode)
        {
            var command = new OdbcCommand($"select part_id from parts where part_name = '{partName}' and part_code = '{partCode}';", connection);
            return command.ExecuteScalar().ToString();
        }
        public bool CheckAutoIn(string autoName, string autoModel, string autoFuel, string autoYear)
        {
            var command = new OdbcCommand($"select 1 from autos where auto_name = '{autoName}' and auto_model = '{autoModel}' and auto_fuelvolume = '{autoFuel}' and auto_year = '{autoYear}';", connection);
            return 1 == Convert.ToInt32(command.ExecuteScalar());
        }
        private void AddPartsInfoButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(PartsAddPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            if (!CheckAutoIn(AutoNameComboBox.Text, AutoModelComboBox.Text, AutoFuelTypeComboBox.Text, AutoYearComboBox.Text))
            {
                AddNewRowStrip($"Не хотите добавить автомобиль?\n {AutoNameComboBox.Text}  {AutoModelComboBox.Text}  {AutoFuelTypeComboBox.Text}  {AutoYearComboBox.Text}?'");
                ThrowError("Выбранного автомобиля нет в базе данных. Повторите выбор.");
                return;
            }
            var insert = new OdbcCommand($"insert into parts values (default, '{PartNameAddTextBox.Text}', '{PartCodeAddTextBox.Text}', {ReturnCategoryId(CategoryComboBox.Text)},  {ReturnAutoId(AutoNameComboBox.Text, AutoModelComboBox.Text, AutoFuelTypeComboBox.Text, AutoYearComboBox.Text)}, '{PartCostAddTextBox.Text}', '{PartDescriptionTextBox.Text}', '{PartAmountTextBox.Text}');", connection);
            insert.ExecuteNonQuery();
            ShowSuccess("Данные в таблицу запчастей успешно добавлены.");
            this.partsTableAdapter.Fill(this.vvgcarpartsDataSet.parts);
        }
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            connection.Close();
        }
        private void AutoNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillByAutoModel(AutoModelComboBox, AutoNameComboBox.Text);
        }
        private void AutoModelComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillByAutoFuel(AutoFuelTypeComboBox, AutoNameComboBox.Text, AutoModelComboBox.Text);
        }
        private void AutoFuelTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillByAutoYear(AutoYearComboBox, AutoNameComboBox.Text, AutoModelComboBox.Text, AutoFuelTypeComboBox.Text);
        }
        private void SellsTableButton_Click(object sender, EventArgs e)
        {
            PartPhotosDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            PartsDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            PartPhotosDataGrid.Focus();
            ConfirmPanel.Visible = false;
            HideAllAdvPanels();
        }
        public string ReturnMonth(string date)
        {
            var str = date.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            return str[1];
        }
        public void FillComboBoxMonth(ComboBox box, DataGridView dgv)
        {
            //box.Items.Clear();
            //for (int i = 0; i < dgv.RowCount; i++)
            //{
            //    if (!box.Items.Contains((ReturnStringMonth(ReturnMonth(PartPhotosDataGrid.Rows[i].Cells[5].Value.ToString())))))
            //        box.Items.Add(ReturnStringMonth(ReturnMonth(PartPhotosDataGrid.Rows[i].Cells[5].Value.ToString())));
            //}
            //box.Items.Add("За всё время");
        }
        public string ReturnStringMonth(string date)
        {
            string monthName = string.Empty;
            switch (date)
            {
                case "01":
                    monthName = "Январь";
                    break;
                case "02":
                    monthName = "Февраль";
                    break;
                case "03":
                    monthName = "Март";
                    break;
                case "04":
                    monthName = "Апрель";
                    break;
                case "05":
                    monthName = "Май";
                    break;
                case "06":
                    monthName = "Июнь";
                    break;
                case "07":
                    monthName = "Июль";
                    break;
                case "08":
                    monthName = "Август";
                    break;
                case "09":
                    monthName = "Сентябрь";
                    break;
                case "10":
                    monthName = "Октябрь";
                    break;
                case "11":
                    monthName = "Ноябрь";
                    break;
                case "12":
                    monthName = "Декабрь";
                    break;
            }
            return monthName;
        }
        public string ReturnNumberMonth(string date)
        {
            string id = string.Empty;
            switch (date)
            {
                case "Январь":
                    id = "1";
                    break;
                case "Февраль":
                    id = "2";
                    break;
                case "Март":
                    id = "3";
                    break;
                case "Апрель":
                    id = "4";
                    break;
                case "Май":
                    id = "5";
                    break;
                case "Июнь":
                    id = "6";
                    break;
                case "Июль":
                    id = "7";
                    break;
                case "Август":
                    id = "8";
                    break;
                case "Сентябрь":
                    id = "9";
                    break;
                case "Октябрь":
                    id = "10";
                    break;
                case "Ноябрь":
                    id = "11";
                    break;
                case "Декабрь":
                    id = "12";
                    break;
            }
            return id;
        }
        private void SellsViewDataGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
        }

        public int ReturnCarFromGrid(string autoName, string autoModel)
        {
            int id = 0;
            for (int i = 0; i < AutoDataGrid.RowCount; i++)
            {
                if (AutoDataGrid.Rows[i].Cells[1].Value.ToString() == autoName)
                {
                    if (AutoDataGrid.Rows[i].Cells[2].Value.ToString() == autoModel)
                    {
                        id = Convert.ToInt32(AutoDataGrid.Rows[i].Cells[0].Value.ToString());
                        break;
                    }
                }
            }
            return id;
        }

        private void PartsDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void AddNewAutoButton_Click(object sender, EventArgs e)
        {
            if (PartsAddPanel.Visible)
            {
                AutoAddPanel.Visible = true;
                AutoNameAddTextBox.Text = AutoNameComboBox.Text;
                AutoModelAddTextBox.Text = AutoModelComboBox.Text;
                AutoFuelTypeAddTextBox.Text = AutoModelComboBox.Text;
                AutoYearAddTextBox.Text = AutoYearComboBox.Text;
                NewRowPanel.Visible = false;
                PartsAddPanel.Visible = false;
                AutoAddPanel.Location = new System.Drawing.Point(546, 106);
                AddInfoButton.Enabled = false;
                DeleteInfoButton.Enabled = false;
                FindInfoButton.Enabled = false;
            }
        }

        private void CancelAddAutoButton_Click(object sender, EventArgs e)
        {
            NewRowPanel.Visible = false;
        }

        private void GoToPartsOfThisCar_Click(object sender, EventArgs e)
        {

            AdvancedDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            var command = new OdbcDataAdapter($"select ID as №, PName as НАЗВАНИЕ, PCode as МАРКИРОВКА, CName as КАТЕГОРИЯ, AName as 'МАРКА АВТО', AModel as 'МОДЕЛЬ АВТО', PCost as ЦЕНА, PAmount as КОЛИЧЕСТВО, PPhoto from PartsView where AName = '{AutoDataGrid.CurrentRow.Cells[1].Value}' and AModel = '{AutoDataGrid.CurrentRow.Cells[2].Value}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "PartsView");

            AdvancedDataGrid.DataSource = ds.Tables[0];
            AdvancedDataGrid.Focus();
            AdvancedDataGrid.Columns[8].Visible = false;
            AdvancedDataGrid.Columns[0].Width = 50;

        }

        private void MainDetailsTableButton_Click(object sender, EventArgs e)
        {
            ReviewsDataGrid.Visible = true;
            AutoDataGrid.Visible = false;
            PartsDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            AdvancedDataGrid.Visible = false;
            ReviewsDataGrid.Focus();
            HideAllAdvPanels();
            this.reviewsTableAdapter.Fill(this.vvgcarpartsDataSet.reviews);
        }

        private void ExportWord_MouseEnter(object sender, EventArgs e)
        {
            WordExportLabel.ForeColor = Color.FromArgb(82, 171, 133);
        }

        private void ExportExcel_MouseEnter(object sender, EventArgs e)
        {
            ExcelExportLabel.ForeColor = Color.FromArgb(82, 171, 133);
        }



        void ExportLabelsHover_MouseEnter(object sender, EventArgs e)
        {

        }

        private void ExcelExportLabel_MouseLeave(object sender, EventArgs e)
        {
            ExcelExportLabel.ForeColor = Color.White;
        }

        private void WordExportLabel_MouseLeave(object sender, EventArgs e)
        {
            WordExportLabel.ForeColor = Color.White;
        }

        private void WordPageExportLabel_MouseLeave(object sender, EventArgs e)
        {

        }

        private void SelectedCarDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            //if (AdvancedDataGrid.Visible && AdvancedDataGrid.CurrentRow != null && AdvancedDataGrid.DataSource != null)
            //{
            //    MainPictureBox.Image = Properties.Resources.load1;
            //    MainPictureBox.ImageLocation = AdvancedDataGrid.CurrentRow.Cells[8].Value.ToString();
            //    CurrentItemLabel.Text = AdvancedDataGrid.CurrentRow.Cells[1].Value.ToString() + " " + AdvancedDataGrid.CurrentRow.Cells[4].Value.ToString() + " " + AdvancedDataGrid.CurrentRow.Cells[5].Value.ToString();
            //}
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PartNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {


            FillByPartCode(PartCodeSellsViewCB, PartNameSellsComboBox.Text);
        }

        private void AddSellsRowButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(PartPhotoAddPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            if (!CheckPartInTable(PartNameSellsComboBox.Text, AutoNameSellsComboBox.Text, AutoModelSellsComboBox.Text, AutoFueltypeSellsComboBox.Text, AutoYearSellsComboBox.Text))
            {
                ThrowError("Данная запчасть отсутствует в базе данных. Повторите ввод.");
                return;
            }

            var command = new OdbcCommand($"insert into part_photos values(default,'{ReturnPartId(PartNameSellsComboBox.Text, ReturnAutoId(AutoNameSellsComboBox.Text, AutoModelSellsComboBox.Text, AutoFueltypeSellsComboBox.Text, AutoYearSellsComboBox.Text), PartCodeSellsViewCB.Text)}', '{PartURLTextBox.Text}');", connection);
            command.ExecuteNonQuery();
            ShowSuccess("Фото для запчасти успешно добавлено");

        }

        public string GetCorrectDate(string date)
        {
            var allDate = date.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            string dat = allDate.Last() + "-" + allDate[allDate.Length - 2] + "-" + allDate[allDate.Length - 3];
            return dat;
        }

        public bool CheckPartInTable(string partName, string autoName, string autoModel, string autoFuel, string autoYear)
        {
            if (!CheckAutoIn(autoName, autoModel, autoFuel, autoYear))
            {
                return false;
            }
            else
            {
                var command = new OdbcCommand($"select 1 from parts where part_name = '{partName}' and car_id = '{ReturnAutoId(autoName, autoModel, autoFuel, autoYear)}'", connection);
                return 1 == Convert.ToInt32(command.ExecuteScalar());
            }
        }

        public virtual bool CheckPartInTable(string partName, string partCode)
        {
            var command = new OdbcCommand($"select 1 from parts where part_name = '{partName}' and part_code = '{partCode}'", connection);
            return 1 == Convert.ToInt32(command.ExecuteScalar());
        }

        private void CancelAddSellsButton_Click(object sender, EventArgs e)
        {
            CancelAddInfo(PartPhotoAddPanel);
            PartPhotoAddPanel.Visible = false;
            OnButtons();
        }

        private void AutoYearAddTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void PartCostAddTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void PartAmountTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void AutoNameSellsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void NewRowPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void AutoNameSellsComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var command = new OdbcDataAdapter($"select distinct auto_model from autos inner join parts on autos.id = parts.car_id where part_name = '{PartNameSellsComboBox.Text}' and auto_name = '{AutoNameSellsComboBox.Text}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "autos");

            AutoModelSellsComboBox.DataSource = ds.Tables[0];
            AutoModelSellsComboBox.DisplayMember = "auto_model";
        }

        private void AutoModelSellsComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var command = new OdbcDataAdapter($"select distinct auto_fuelvolume from autos inner join parts on autos.id = parts.car_id where part_name = '{PartNameSellsComboBox.Text}' and auto_name = '{AutoNameSellsComboBox.Text}' and auto_model = '{AutoModelSellsComboBox.Text}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "autos");

            AutoFueltypeSellsComboBox.DataSource = ds.Tables[0];
            AutoFueltypeSellsComboBox.DisplayMember = "auto_fuelvolume";
        }

        private void AutoFueltypeSellsComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var command = new OdbcDataAdapter($"select distinct auto_year from autos inner join parts on autos.id = parts.car_id where part_name = '{PartNameSellsComboBox.Text}' and auto_name = '{AutoNameSellsComboBox.Text}' and auto_model = '{AutoModelSellsComboBox.Text}' and auto_fuelvolume = '{AutoFueltypeSellsComboBox.Text}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "autos");

            AutoYearSellsComboBox.DataSource = ds.Tables[0];
            AutoYearSellsComboBox.DisplayMember = "auto_year";
        }

        private void AutoFueltypeSellsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SellsViewDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DeleteInfoButton_Click(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible)
            {
                ConfirmMove("Вы действительно хотите удалить\n" +
                            $"{AutoDataGrid.CurrentRow.Cells[1].Value.ToString()} " +
                            $"{AutoDataGrid.CurrentRow.Cells[2].Value.ToString()} " +
                            $"{AutoDataGrid.CurrentRow.Cells[3].Value.ToString()} " +
                            $"{AutoDataGrid.CurrentRow.Cells[4].Value.ToString()} " +
                            $"{AutoDataGrid.CurrentRow.Cells[6].Value.ToString()}?");
                AutoDataGrid.Enabled = false;
                OffButtons();

            }
            if (PartsDataGrid.Visible)
            {
                ConfirmMove("Вы действительно хотите удалить\n" +
                            $"{PartsDataGrid.CurrentRow.Cells[1].Value.ToString()} " +
                            $"{PartsDataGrid.CurrentRow.Cells[2].Value.ToString()} " +
                            $"{PartsDataGrid.CurrentRow.Cells[3].Value.ToString()} " +
                            $"{PartsDataGrid.CurrentRow.Cells[4].Value.ToString()} " +
                            $"{PartsDataGrid.CurrentRow.Cells[5].Value.ToString()} " +
                            $"{PartsDataGrid.CurrentRow.Cells[6].Value.ToString()}?");
                PartsDataGrid.Enabled = false;
                OffButtons();
            }
            if (PartPhotosDataGrid.Visible)
            {
                ConfirmMove("Вы действительно хотите удалить\n" +
                            $"{PartPhotosDataGrid.CurrentRow.Cells[1].Value.ToString()} " +
                            $"{PartPhotosDataGrid.CurrentRow.Cells[2].Value.ToString()} " +
                            $"{PartPhotosDataGrid.CurrentRow.Cells[3].Value.ToString()} " +
                            $"{PartPhotosDataGrid.CurrentRow.Cells[4].Value.ToString()} " +
                            $"{PartPhotosDataGrid.CurrentRow.Cells[7].Value.ToString()}?");
                PartPhotosDataGrid.Enabled = false;
                OffButtons();
            }
            //if (ReviewsDataGrid.Visible)
            //{
            //    ConfirmMove("Вы действительно хотите удалить\n" +
            //                $"{ReviewsDataGrid.CurrentRow.Cells[1].Value.ToString()} " +
            //                $"{ReviewsDataGrid.CurrentRow.Cells[2].Value.ToString()} " +
            //                $"{ReviewsDataGrid.CurrentRow.Cells[4].Value.ToString()} " +
            //                $"{ReviewsDataGrid.CurrentRow.Cells[5].Value.ToString()}?");
            //    ReviewsDataGrid.Enabled = false;
            //    OffButtons();
            //}
            if (CategoriesDataGrid.Visible)
            {
                ConfirmMove("Вы действительно хотите удалить\n" +
                            $"{CategoriesDataGrid.CurrentRow.Cells[1].Value.ToString()}");
                CategoriesDataGrid.Enabled = false;
                OffButtons();
            }
            if (UsersDataGrid.Visible)
            {
                ConfirmMove("Вы действительно хотите удалить\n" +
                    $"пользователя {UsersDataGrid.CurrentRow.Cells[1].Value.ToString()}?");
                UsersDataGrid.Enabled = false;
                OffButtons();
            }
        }

        public void HideAllAdvPanels()
        {
            OnButtons();
            for (int i = 0; i < this.Controls.Count; i++)
            {
                if (this.Controls[i].GetType() == typeof(Panel) && this.Controls[i].Name != "HidingPanel")
                {
                    this.Controls[i].Visible = false;
                }
            }
        }

        private void CancelDeleteAutoButton_Click(object sender, EventArgs e)
        {
            DeleteAutoPanel.Visible = false;
            OnButtons();
        }

        private void DeleteAutoButton_Click(object sender, EventArgs e)
        {
            if (!CheckAutoIn(DANComboBox.Text, DAMComboBox.Text, DAFComboBox.Text, DAYComboBox.Text))
            {
                ThrowError("Выбранного автомобиля нет в базе данных. Повторите выбор.");
                return;
            }

            for (int i = 0; i < DeleteAutoPanel.Controls.Count; i++)
            {
                if (DeleteAutoPanel.Controls[i].GetType() == typeof(ComboBox) && DeleteAutoPanel.Controls[i].Text == string.Empty)
                {
                    ThrowError("Один из выпадающих списков пуст. Повторите ввод.");
                    return;
                }
            }

            ConfirmMove("Вы действительно хотите удалить автомобиль? Удалив его, вы потеряете всю информацию о нем в других таблицах.");
        }

        private void DANComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoModel(DAMComboBox, DANComboBox.Text);

        }

        private void DAMComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoFuel(DAFComboBox, DANComboBox.Text, DAMComboBox.Text);

        }

        private void DAFComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoYear(DAYComboBox, DANComboBox.Text, DAMComboBox.Text, DAFComboBox.Text);
        }

        private void CancelDeleteButton_Click(object sender, EventArgs e)
        {
            ConfirmPanel.Visible = false;
            AutoDataGrid.Enabled = true;
            PartsDataGrid.Enabled = true;
            PartPhotosDataGrid.Enabled = true;
            OnButtons();
        }

        private void ConfirmDeleteButton_Click(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible)
            {
                if (AutoDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from autos where id = '{AutoDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess($"Авто под номером {AutoDataGrid.CurrentRow.Cells[0].Value.ToString()} успешно удалено из базы данных.");
                    command.ExecuteNonQuery();
                    this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
                    AutoDataGrid.Enabled = true;
                }
            }
            if (PartsDataGrid.Visible)
            {
                if (PartsDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from parts where part_id = '{PartsDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess($"Запчасть под номером {PartsDataGrid.CurrentRow.Cells[0].Value.ToString()} успешно удалена из базы данных.");
                    command.ExecuteNonQuery();
                    this.partsTableAdapter.Fill(this.vvgcarpartsDataSet.parts);
                    ConfirmPanel.Visible = false;
                    PartsDataGrid.Enabled = true;
                }
            }
            if (PartPhotosDataGrid.Visible)
            {
                if (PartPhotosDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from sells where part_id = '{PartPhotosDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess($"Продажа под номером {PartPhotosDataGrid.CurrentRow.Cells[0].Value.ToString()} успешно удалена из базы данных.");
                    command.ExecuteNonQuery();
                    ConfirmPanel.Visible = false;
                    PartPhotosDataGrid.Enabled = true;
                }
            }
            if (ReviewsDataGrid.Visible)
            {
                if (ReviewsDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from maindetails where maindetail_id = '{ReviewsDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess($"Продажа под номером {ReviewsDataGrid.CurrentRow.Cells[0].Value.ToString()} успешно удалена из базы данных.");
                    command.ExecuteNonQuery();
                    ConfirmPanel.Visible = false;
                    PartPhotosDataGrid.Enabled = true;
                }
            }
            if (CategoriesDataGrid.Visible)
            {
                if (CategoriesDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from categories where id = '{CategoriesDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess("Категория успешно удалена.");
                    command.ExecuteNonQuery();
                    this.categoriesTableAdapter.Fill(this.vvgcarpartsDataSet.categories);
                    ConfirmPanel.Visible = false;
                    CategoriesDataGrid.Enabled = true;
                }
            }
            if (UsersDataGrid.Visible)
            {
                if (UsersDataGrid.CurrentRow != null)
                {
                    var command = new OdbcCommand($"delete from users where userId = '{UsersDataGrid.CurrentRow.Cells[0].Value.ToString()}'", connection);
                    ShowSuccess("Пользователь успешно удален.");
                    command.ExecuteNonQuery();
                    this.usersTableAdapter.Fill(this.vvgcarpartsDataSet.users);
                    ConfirmPanel.Visible = false;
                    UsersDataGrid.Enabled = true;
                }
            }
            for (int i = 0; i < this.Controls.Count; i++)
            {
                if (this.Controls[i].GetType() == typeof(DataGridView))
                    this.Controls[i].Enabled = true;
            }
            OnButtons();
        }

        int iFormX, iFormY, iMouseX, iMouseY;

        public void ChangeCoordinates(Panel panel)
        {
            iFormX = panel.Location.X;
            iFormY = panel.Location.Y;
            iMouseX = MousePosition.X;
            iMouseY = MousePosition.Y;
        }

        private void ConfirmPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                ConfirmPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void ErrorPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(ErrorPanel);
        }

        private void ErrorPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                ErrorPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void SuccessPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(SuccessPanel);
        }

        private void SuccessPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                SuccessPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void AutoAddPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                AutoAddPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void AutoAddPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(AutoAddPanel);
        }

        private void PartsAddPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(PartsAddPanel);
        }

        private void PartsAddPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                PartsAddPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void SellsAddPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(PartPhotoAddPanel);
        }

        private void SellsAddPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                PartPhotoAddPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void PNDComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoName(PartAutoNameDeleteCB, PartNameDeleteCB.Text);
        }

        private void CancelPartDeleteButton_Click(object sender, EventArgs e)
        {
            DeletePartPanel.Visible = false;
            OnButtons();
        }

        private void PNDComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void PartAutoNameCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoModelForParts(PartsAutoModelCB, PartNameDeleteCB.Text);
        }

        private void PAMComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillByAutoFuel(PartsFuelAutoCB, PartNameDeleteCB.Text);
        }

        private void PAFComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoYear(PartsAutoYearCB, PartNameDeleteCB.Text);
        }

        private void PAYComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByPartDescription(PartDescripComboBox, PartNameDeleteCB.Text);
        }

        private void PAFComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PAYComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PAYComboBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void PartDeleteButton_Click(object sender, EventArgs e)
        {

            if (!CheckPartInTable(PartNameDeleteCB.Text, PartAutoNameDeleteCB.Text, PartsAutoModelCB.Text, PartsFuelAutoCB.Text, PartsAutoYearCB.Text))
            {
                ThrowError("Выбранной запчасти нет в базе данных. Повторите выбор.");
                return;
            }

            for (int i = 0; i < DeletePartPanel.Controls.Count; i++)
            {
                if (DeletePartPanel.Controls[i].GetType() == typeof(ComboBox) && DeletePartPanel.Controls[i].Text == string.Empty)
                {
                    ThrowError("Один из выпадающих списков пуст. Повторите ввод.");
                    return;
                }
            }

            ConfirmMove("Вы действительно хотите удалить запчасть? Удалив его, вы потеряете всю информацию о нем в других таблицах.");
        }

        private void EditRowsButton_Click(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible)
            {
                if (AutoDataGrid.CurrentRow != null)
                {
                    EditAutoPanel.Visible = true;
                    EditAutoPanel.Location = new System.Drawing.Point(339, 344);
                    EditAutoIdLabel.Text = AutoDataGrid.CurrentRow.Cells[0].Value.ToString();
                    EditAutoNameTextBox.Text = AutoDataGrid.CurrentRow.Cells[1].Value.ToString();
                    EditAutoModelTextBox.Text = AutoDataGrid.CurrentRow.Cells[2].Value.ToString();
                    EditAutoBodyTextBox.Text = AutoDataGrid.CurrentRow.Cells[3].Value.ToString();
                    EditAutoFuelTextBox.Text = AutoDataGrid.CurrentRow.Cells[4].Value.ToString();
                    EditAutoFuelVolumeTextBox.Text = AutoDataGrid.CurrentRow.Cells[5].Value.ToString();
                    EditAutoTransmTextBox.Text = AutoDataGrid.CurrentRow.Cells[6].Value.ToString();
                    EditAutoYearTextBox.Text = AutoDataGrid.CurrentRow.Cells[7].Value.ToString();
                    EditAutoDriveTextBox.Text = AutoDataGrid.CurrentRow.Cells[8].Value.ToString();
                    EditAutoPhotoTextBox.Text = AutoDataGrid.CurrentRow.Cells[9].Value.ToString();
                    OffButtons();
                    EditAutoPanel.BringToFront();
                }
                else
                {
                    ThrowError("В таблице авто нет строк для редактирования.");
                    return;
                }
            }
            if (PartsDataGrid.Visible)
            {
                if (PartsDataGrid.CurrentRow != null)
                {
                    EditPartPanel.Visible = true;
                    EditPartPanel.Location = new System.Drawing.Point(339, 344);
                    FillByAutoName(EditPartAutoNameCB);
                    FillByCategoryName(EditPartCategoryCB);
                    EditPartIdLabelRO.Text = PartsDataGrid.CurrentRow.Cells[0].Value.ToString();
                    EditPartNameTB.Text = PartsDataGrid.CurrentRow.Cells[1].Value.ToString();
                    EditPartCodeTB.Text = PartsDataGrid.CurrentRow.Cells[2].Value.ToString();
                    EditPartCategoryCB.Text = GetCategoryName(PartsDataGrid.CurrentRow.Cells[3].Value.ToString());
                    FillByAutoNameByID(EditPartAutoNameCB, PartsDataGrid.CurrentRow.Cells[4].Value.ToString());
                    EditPartCostTB.Text = PartsDataGrid.CurrentRow.Cells[5].Value.ToString();
                    EditPartDescTB.Text = PartsDataGrid.CurrentRow.Cells[6].Value.ToString();
                    EditPartAmountTB.Text = PartsDataGrid.CurrentRow.Cells[7].Value.ToString();
                    OffButtons();
                    EditPartPanel.BringToFront();
                }
                else
                {
                    ThrowError("В таблице запчастей нет строк для редактирования.");
                    return;
                }
            }
            if (PartPhotosDataGrid.Visible)
            {
                if (PartPhotosDataGrid.CurrentRow != null)
                {
                    EditSellsPanel.Visible = true;
                    EditSellsPanel.Location = new System.Drawing.Point(339, 344);
                    EditSellIdLabel.Text = PartPhotosDataGrid.CurrentRow.Cells[0].Value.ToString();
                    ReviewerNameEdit.Text = PartPhotosDataGrid.CurrentRow.Cells[1].Value.ToString();
                    ReviewTextEdit.Text = PartPhotosDataGrid.CurrentRow.Cells[2].Value.ToString();
                    EditReviewDate.Text = PartPhotosDataGrid.CurrentRow.Cells[3].Value.ToString();
                    ReviewCodeEdit.Text = PartPhotosDataGrid.CurrentRow.Cells[4].Value.ToString();
                    OffButtons();
                    EditSellsPanel.BringToFront();
                }
                else
                {
                    ThrowError("В таблице фотографий нет строк для редактирования.");
                    return;
                }

            }
            if (ReviewsDataGrid.Visible)
            {
                if (ReviewsDataGrid.CurrentRow != null)
                {

                    EditSellsPanel.Visible = true;
                    EditSellsPanel.Location = new System.Drawing.Point(339, 344);
                    EditSellIdLabel.Text = ReviewsDataGrid.CurrentRow.Cells[0].Value.ToString();
                    ReviewerNameEdit.Text = ReviewsDataGrid.CurrentRow.Cells[1].Value.ToString();
                    ReviewTextEdit.Text = ReviewsDataGrid.CurrentRow.Cells[2].Value.ToString();
                    EditReviewDatePicker.Text = ReviewsDataGrid.CurrentRow.Cells[3].Value.ToString();
                    ReviewCodeEdit.Text = ReviewsDataGrid.CurrentRow.Cells[4].Value.ToString();
                    OffButtons();
                }
                else
                {
                    ThrowError("В таблице отзывов нет строк для редактирования.");
                    return;
                }
            }

        }

        private void EditAutoPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void EditAutoYearTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void EditAutoYearTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void CancelEditAutoButton_Click(object sender, EventArgs e)
        {
            EditAutoPanel.Visible = false;
            OnButtons();
        }

        private void EditAutoButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(EditAutoPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            var command = new OdbcCommand($"update autos set auto_name = '{EditAutoNameTextBox.Text}', auto_model = '{EditAutoModelTextBox.Text}', auto_bodytype = '{EditAutoBodyTextBox.Text}', auto_fueltype = '{EditAutoFuelTextBox.Text}', auto_fuelvolume = '{EditAutoFuelVolumeTextBox.Text}', auto_transmissiotype = '{EditAutoTransmTextBox.Text}', " +
                $" auto_year = '{EditAutoYearTextBox.Text}', auto_drivetype = '{EditAutoDriveTextBox.Text}', auto_photo = '{EditAutoPhotoTextBox.Text}' where id = '{EditAutoIdLabel.Text}';", connection);
            command.ExecuteNonQuery();
            ShowSuccess($"Запись под номером {EditAutoIdLabel.Text} в таблице авто успешно обновлена.");
            this.autosTableAdapter.Fill(this.vvgcarpartsDataSet.autos);
        }

        private void EditAutoPanel_MouseDown(object sender, MouseEventArgs e)
        {
            iFormX = EditAutoPanel.Location.X;
            iFormY = EditAutoPanel.Location.Y;
            iMouseX = MousePosition.X;
            iMouseY = MousePosition.Y;
        }

        private void EditAutoPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                EditAutoPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void EditPartAutoNameCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoModel(EditPartAutoModelCB, EditPartAutoNameCB.Text);
        }

        private void EditPartAutoModelCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoFuel(EditPartAutoFuelCB, EditPartAutoNameCB.Text, EditPartAutoModelCB.Text);
        }

        private void EditPartAutoFuelCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoYear(EditPartAutoYearCB, EditPartAutoNameCB.Text, EditPartAutoModelCB.Text, EditPartAutoFuelCB.Text);
        }

        private void EditPartPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(EditPartPanel);
        }

        private void EditPartPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                EditPartPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void CancelEditPartButton_Click(object sender, EventArgs e)
        {
            EditPartPanel.Visible = false;
            OnButtons();
        }

        private void EditPartButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(EditPartPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            if (!CheckAutoIn(EditPartAutoNameCB.Text, EditPartAutoModelCB.Text, EditPartAutoFuelCB.Text, EditPartAutoYearCB.Text))
            {
                ThrowError("Выбранного автомобиля нет в базе данных. Повторите выбор.");
                return;
            }
            var command = new OdbcCommand($"update parts set part_name = '{EditPartNameTB.Text}', part_code = '{EditPartCodeTB.Text}', cat_id = '{ReturnCategoryId(EditPartCategoryCB.Text)}', car_id = '{ReturnAutoId(EditPartAutoNameCB.Text, EditPartAutoModelCB.Text, EditPartAutoFuelCB.Text, EditPartAutoYearCB.Text)}'," +
                $"part_cost = '{EditPartCostTB.Text}', part_description = '{EditPartDescTB.Text}', part_amount = '{EditPartAmountTB.Text}' where part_id = '{EditPartIdLabelRO.Text}';", connection);
            command.ExecuteNonQuery();
            ShowSuccess($"Запись под номером {EditPartIdLabelRO.Text} в таблице запчастей успешно обновлена.");
            this.partsTableAdapter.Fill(this.vvgcarpartsDataSet.parts);
        }

        private void EditPartAmountTB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void CancelEditSellsButton_Click(object sender, EventArgs e)
        {
            EditSellsPanel.Visible = false;
            OnButtons();
        }

        private void EditSellsButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(EditSellsPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }

            var command = new OdbcCommand($"update reviews set reviewer_name = '{ReviewerNameEdit.Text}', reviewer_text = '{ReviewTextEdit.Text}', reviewer_date = '{GetCorrectDate(EditReviewDatePicker.Text)}', mark = '{ReviewCodeEdit.Text}' where review_id = {EditSellIdLabel.Text};", connection);
            command.ExecuteNonQuery();
            ShowSuccess($"Запись в таблице отзывов под номером {EditSellIdLabel.Text} успешно обновлена.");
            this.reviewsTableAdapter.Fill(this.vvgcarpartsDataSet.reviews);
        }


        private void EditSellsAutoNameCB_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void EditSellsAutoModelCB_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void EditSellsAutoFuelCB_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void EditSellsEarningTB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void EarningSellTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void EditSellsPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(EditSellsPanel);
        }

        private void EditSellsPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                EditSellsPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }


        private void FindInfoButton_Click(object sender, EventArgs e)
        {
            OffButtons();
            TablesComboBox.Text = "Автомобили";
            AdvancedDataGrid.Visible = true;
            PartsDataGrid.Visible = false;
            AutoDataGrid.Visible = false;
            UsersDataGrid.Visible = false;
            CategoriesDataGrid.Visible = false;
            ReviewsDataGrid.Visible = false;
            PartPhotosDataGrid.Visible = false;
            FindInfoPanel.Visible = true;
            FindInfoPanel.Location = new System.Drawing.Point(671, 739);
            FindInfoPanel.BringToFront();
        }

        private void CancelFindInfoButton_Click(object sender, EventArgs e)
        {
            FindInfoPanel.Visible = false;
            OnButtons();
            AdvancedDataGrid.Visible = false;
            AutoDataGrid.Visible = true;

        }

        private void FindInfoTB_TextChanged(object sender, EventArgs e)
        {
            if (FindInfoTB.Text.Length > 0)
            {
                if (TablesComboBox.Text == "Автомобили")
                {
                    var command = new OdbcDataAdapter($"select id as '№', auto_name as 'МАРКА АВТО', auto_fuelvolume as 'ОБЪЁМ', auto_fuelvolume as 'ТИП ТОПЛИВА', auto_bodytype as 'ТИП КУЗОВА',auto_model as 'МОДЕЛЬ АВТО' , auto_transmissiotype as 'ТИП КПП', auto_drivetype as 'ПРИВОД', auto_year as 'ГОД ВЫПУСКА', auto_photo from autos where concat(auto_name,auto_model, auto_fuelvolume, " +
                        $" auto_bodytype, auto_transmissiotype, auto_drivetype, auto_year) like '%{FindInfoTB.Text}%';", connection);
                    var columns = new OdbcCommandBuilder(command);
                    DataSet ds = new DataSet();
                    command.Fill(ds, "autos");
                    AdvancedDataGrid.DataSource = ds.Tables[0];
                    AdvancedDataGrid.Columns[0].Width = 50;
                    AdvancedDataGrid.Columns[8].Visible = false;
                }
                if (TablesComboBox.Text == "Запчасти")
                {
                    var command = new OdbcDataAdapter($"select part_id as '№', part_name as 'НАЗВАНИЕ', part_code as 'МАРКИРОВКА', cat_id as 'ID КАТЕГОРИИ', car_id as 'ID АВТО', part_cost as '', part_description as '', part_amount as '' from parts " +
                        $" where concat (part_name, part_code, part_code, cat_id, car_id, part_cost, part_description, part_amount) like '%{FindInfoTB.Text}%';", connection);
                    var columns = new OdbcCommandBuilder(command);
                    DataSet ds = new DataSet();
                    command.Fill(ds, "parts");
                    AdvancedDataGrid.DataSource = ds.Tables[0];
                    AdvancedDataGrid.Columns[0].Width = 50;
                }
                if (TablesComboBox.Text == "Фото")
                {
                    var command = new OdbcDataAdapter($"select photo_id as '№', part_id as 'ID ЗАПЧАСТИ', photo_url as 'ФОТО URL' from part_photos where concat(photo_id, part_id, photo_url) like '%{FindInfoTB.Text}%';", connection);
                    var columns = new OdbcCommandBuilder(command);
                    DataSet ds = new DataSet();
                    command.Fill(ds, "part_photos");
                    AdvancedDataGrid.DataSource = ds.Tables[0];
                    AdvancedDataGrid.Columns[0].Width = 50;

                }
                if (TablesComboBox.Text == "Отзывы")
                {
                    AdvancedDataGrid.Columns.Clear();
                    var command = new OdbcDataAdapter($"select review_id as '№', reviewer_name as 'ИМЯ ОТЗЫВАТЕЛЯ', reviewer_text as 'ТЕКСТ ОТЗЫВА', reviewer_date as 'ДАТА', mark as 'МОДЕРАЦИЯ' from reviews where concat(review_id, " +
                        $" reviewer_name, reviewer_text, reviewer_date, mark) like '%{FindInfoTB.Text}%';", connection);
                    var columns = new OdbcCommandBuilder(command);
                    DataSet ds = new DataSet();
                    command.Fill(ds, "reviews");
                    AdvancedDataGrid.DataSource = ds.Tables[0];
                    AdvancedDataGrid.Columns[0].Width = 50;

                }
            }
        }

        private void ExcelExportLabel_Click(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible)
                ExportExcel(AutoDataGrid);
            if (PartsDataGrid.Visible)
                ExportExcel(PartsDataGrid);
            if (PartPhotosDataGrid.Visible)
                ExportExcel(PartPhotosDataGrid);
            if (ReviewsDataGrid.Visible)
                ExportExcel(ReviewsDataGrid);
            if (AdvancedDataGrid.Visible)
                ExportExcel(AdvancedDataGrid);
        }

        public void ExportExcel(DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    if (dgv.Rows[i].Cells[j].GetType().Name == "DataGridViewComboBoxCell")
                    {
                        DataGridViewComboBoxCell dgvcbc = new DataGridViewComboBoxCell();
                        dgvcbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[j];
                        ExcelApp.Cells[i + 2, j + 1] = dgvcbc.EditedFormattedValue;
                    }
                    else
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value;
                    }
                }
                ExcelWorkSheet.Rows[i + 1].Style.Font.Size = 14;
                if (i % 2 == 0)
                    ExcelWorkSheet.Rows[i + 2].Interior.Color = Color.FromArgb(37, 51, 64);
                else
                    ExcelWorkSheet.Rows[i + 2].Interior.Color = Color.FromArgb(43, 57, 72);
            }
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                ExcelApp.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
                ExcelApp.Cells[1, i + 1].Font.Size = 17;
                ExcelWorkSheet.Cells[1, i + 1].EntireRow.Font.Bold = true;
                ExcelWorkSheet.Cells[1, i + 1].Interior.Color = Color.FromArgb(31, 176, 137);
            }
            ExcelApp.Columns.ColumnWidth = 20;
            ExcelApp.StandardFont = "Century Gothic";
            ExcelWorkSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ExcelWorkSheet.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ExcelWorkSheet.Cells.ColumnWidth = 40;
            ExcelWorkSheet.Cells[1].ColumnWidth = 10;
            ExcelWorkSheet.Cells[9].ColumnWidth = 140;
            ExcelWorkSheet.Cells.Font.Color = Color.White;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        public void ExportMonthExcel(DataGridView dgv, string text)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    if (dgv.Rows[i].Cells[j].GetType().Name == "DataGridViewComboBoxCell")
                    {
                        DataGridViewComboBoxCell dgvcbc = new DataGridViewComboBoxCell();
                        dgvcbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[j];
                        ExcelApp.Cells[i + 2, j + 1] = dgvcbc.EditedFormattedValue;
                    }
                    else
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value;
                    }
                    ExcelWorkSheet.Cells[i + 1, j + 1].Style.Font.Size = 14;
                    if (i % 2 == 0)
                        ExcelWorkSheet.Cells[i + 2, j + 1].Interior.Color = Color.FromArgb(37, 51, 64);
                    else
                        ExcelWorkSheet.Cells[i + 2, j + 1].Interior.Color = Color.FromArgb(43, 57, 72);
                }
            }
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                ExcelApp.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
                ExcelApp.Cells[1, i + 1].Font.Size = 17;
                ExcelWorkSheet.Cells[1, i + 1].EntireRow.Font.Bold = true;
                ExcelWorkSheet.Cells[1, i + 1].Interior.Color = Color.FromArgb(31, 176, 137);
            }


            Microsoft.Office.Interop.Excel.Range excelRange = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.get_Range($"A{dgv.RowCount + 2}", $"B{dgv.RowCount + 2}");
            excelRange.Merge(Type.Missing);
            Microsoft.Office.Interop.Excel.Range excelRange2 = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.get_Range($"E{dgv.RowCount + 2}", $"F{dgv.RowCount + 2}");
            excelRange2.Merge(Type.Missing);

            ExcelApp.Cells[dgv.RowCount + 2, 1] = text;
            ExcelApp.Cells[dgv.RowCount + 2, 3] = ReturnSum(AdvancedDataGrid, 7) + " BYN";
            ExcelApp.Cells[dgv.RowCount + 2, 5] = "КОЛИЧЕСТВО ПРОДАЖ";
            ExcelApp.Cells[dgv.RowCount + 2, 7] = dgv.RowCount;
            ExcelApp.Cells[dgv.RowCount + 2, 1].Interior.Color = Color.FromArgb(30, 119, 176);
            ExcelApp.Cells[dgv.RowCount + 2, 3].Interior.Color = Color.FromArgb(30, 119, 176);
            ExcelApp.Cells[dgv.RowCount + 2, 5].Interior.Color = Color.FromArgb(30, 119, 176);
            ExcelApp.Cells[dgv.RowCount + 2, 7].Interior.Color = Color.FromArgb(30, 119, 176);
            ExcelApp.Columns.ColumnWidth = 20;
            ExcelApp.StandardFont = "Century Gothic";
            ExcelWorkSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ExcelWorkSheet.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ExcelWorkSheet.Cells.ColumnWidth = 20;
            ExcelWorkSheet.Cells[1].ColumnWidth = 10;
            ExcelWorkSheet.Cells.Font.Color = Color.White;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        public int ReturnSum(DataGridView dgv, int cells)
        {
            var sum = new List<int>();
            for (int i = 0; i < dgv.RowCount; i++)
            {
                var str = dgv.Rows[i].Cells[cells].Value.ToString();
                var nums = str.Split(' ');
                sum.Add(int.Parse(nums[0]));
            }
            return sum.Sum();
        }

        public void ExportWord(DataGridView DGV, string tableName)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count - 1;
                int ColumnCount = DGV.Columns.Count - 1;
                object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    }
                }

                Microsoft.Office.Interop.Word.Document oDoc = new Document();

                oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }
                oRange.Text = oTemp;

                object Separator = WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Century Gothic";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                int rgbColorHeader = Information.RGB(31, 176, 137);
                Microsoft.Office.Interop.Word.WdColor headerColor = (Microsoft.Office.Interop.Word.WdColor)rgbColorHeader;

                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                    oDoc.Tables[1].Cell(1, c + 1).Range.Shading.BackgroundPatternColor = headerColor;
                    oDoc.Tables[1].Cell(1, c + 1).Range.Font.Color = WdColor.wdColorWhite;
                    oDoc.Tables[1].Cell(1, c + 1).Range.Font.Size = 15;
                    oDoc.Tables[1].Cell(1, c + 1).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    oDoc.Tables[1].Cell(1, c + 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                int rgbColor = Information.RGB(37, 51, 64);
                Microsoft.Office.Interop.Word.WdColor mainColor = (Microsoft.Office.Interop.Word.WdColor)rgbColor;

                int rgbColor1 = Information.RGB(43, 57, 72);
                Microsoft.Office.Interop.Word.WdColor advColor = (Microsoft.Office.Interop.Word.WdColor)rgbColor1;

                for (int i = 0; i <= RowCount; i++)
                {
                    oDoc.Tables[1].Rows[i + 2].Range.Font.Name = "Century Gothic";
                    oDoc.Tables[1].Rows[i + 2].Range.Font.Size = 13;
                    oDoc.Tables[1].Rows[i + 2].Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    oDoc.Tables[1].Rows[i + 2].Range.Font.Color = WdColor.wdColorWhite;
                    oDoc.Tables[1].Rows[i + 2].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    if (i % 2 == 0)
                        oDoc.Tables[1].Rows[i + 2].Range.Shading.BackgroundPatternColor = mainColor;
                    else
                        oDoc.Tables[1].Rows[i + 2].Range.Shading.BackgroundPatternColor = advColor;

                }

                oDoc.Application.Visible = true;
            }
        }

        private void WordExportLabel_Click(object sender, EventArgs e)
        {
            if (AutoDataGrid.Visible)
                ExportWord(AutoDataGrid, "Автомобили");
            if (PartsDataGrid.Visible)
                ExportWord(PartsDataGrid, "Запчасти");
            if (PartPhotosDataGrid.Visible)
                ExportWord(PartPhotosDataGrid, "Продажи");
            if (AdvancedDataGrid.Visible)
                ExportWord(AdvancedDataGrid, "Поиск");
            if (ReviewsDataGrid.Visible)
                ExportWord(ReviewsDataGrid, "Одни запчасти на разные автомобили");
        }

        private void DetailRowAddButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(DetailAddPanel))
            {
                ThrowError("Одно из полей для добавления пусто. Повторите ввод.");
                return;
            }
            if (!CheckPartInTable(DetailNameCB.Text, DetailPartCodeCB.Text))
            {
                ThrowError("Выбранная запчасть отсутствует в базе данных. Повторите ввод.");
                return;
            }
            if (!CheckAutoIn(DetailAutoNameCB.Text, DetailAutoModelCB.Text, DetailAutoFuelCB.Text, DetailAutoYearCB.Text))
            {
                ThrowError("Выбранное авто отсутствует в базе данных. Повторите ввод.");
                return;
            }
            var command = new OdbcCommand($"insert into maindetails values(default,'{ReturnPartId(DetailNameCB.Text, DetailPartCodeCB.Text)}', '{ReturnAutoId(DetailAutoNameCB.Text, DetailAutoModelCB.Text, DetailAutoFuelCB.Text, DetailAutoYearCB.Text)}');", connection);
            command.ExecuteNonQuery();
            ShowSuccess("Данные в таблицу деталей на разные автомобили успешно добавлены.");
        }

        private void CancelAddDetailButton_Click(object sender, EventArgs e)
        {
            DetailAddPanel.Visible = false;
            OnButtons();
            CancelAddInfo(DetailAddPanel);
        }

        private void DetailNameCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByPartCode(DetailPartCodeCB, DetailNameCB.Text);
            FillByAutoName(DetailAutoNameCB);
        }

        private void DetailAutoNameCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoModel(DetailAutoModelCB, DetailAutoNameCB.Text);
        }

        private void AutoNameSellsComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void DetailAutoModelCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoFuel(DetailAutoFuelCB, DetailAutoNameCB.Text, DetailAutoModelCB.Text);
        }

        private void DetailAutoFuelCB_SelectedValueChanged(object sender, EventArgs e)
        {
            FillByAutoYear(DetailAutoYearCB, DetailAutoNameCB.Text, DetailAutoModelCB.Text, DetailAutoFuelCB.Text);
        }

        private void DetailsViewDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            //if (ReviewsDataGrid.Visible && ReviewsDataGrid.CurrentRow != null && ReviewsDataGrid.DataSource != null)
            //{
            //    MainPictureBox.Image = Properties.Resources.load1;
            //    MainPictureBox.ImageLocation = ReviewsDataGrid.CurrentRow.Cells[8].Value.ToString();
            //    CurrentItemLabel.Text = ReviewsDataGrid.CurrentRow.Cells[1].Value.ToString() + " " + ReviewsDataGrid.CurrentRow.Cells[2].Value.ToString() + " " + ReviewsDataGrid.CurrentRow.Cells[4].Value.ToString() + " " + ReviewsDataGrid.CurrentRow.Cells[5].Value.ToString();
            //}
        }

        private void EditSellsPartCB_SelectedValueChanged(object sender, EventArgs e)
        {
        }
        public void ThrowICFError(string message)
        {
            InternetConnectionFailedPanel.Visible = true;
            InternetConnectionFailedPanel.Location = new System.Drawing.Point(667, 408);
            ICFRichBox.Text = message;
        }

        private void WordPageExportLabel_Click(object sender, EventArgs e)
        {

        }
        public void ExportPartToWordPage(DataGridView dgv)
        {
            Microsoft.Office.Interop.Word.Application WordApplication = new Microsoft.Office.Interop.Word.Application();
            object fileName = "C:\\Users\\Lenovo\\Documents\\part.docx";
            Microsoft.Office.Interop.Word.Document WordDocument = WordApplication.Documents.Open(fileName);
            WordDocument.Variables["PartName"].Value = dgv.CurrentRow.Cells[1].Value.ToString();
            WordDocument.Variables["PartCode"].Value = dgv.CurrentRow.Cells[2].Value.ToString();
            WordDocument.Variables["Category"].Value = dgv.CurrentRow.Cells[3].Value.ToString();
            WordDocument.Variables["AutoName"].Value = dgv.CurrentRow.Cells[4].Value.ToString();
            WordDocument.Variables["AutoModel"].Value = dgv.CurrentRow.Cells[5].Value.ToString();
            WordDocument.Variables["PartCost"].Value = dgv.CurrentRow.Cells[6].Value.ToString();
            WordDocument.Variables["Description"].Value = dgv.CurrentRow.Cells[7].Value.ToString();
            WordDocument.Variables["PartAmount"].Value = dgv.CurrentRow.Cells[9].Value.ToString();
            Microsoft.Office.Interop.Word.Range docRange = WordDocument.Range();
            Microsoft.Office.Interop.Word.Shape newShape = WordDocument.Shapes[1];
            newShape.Fill.UserPicture(dgv.CurrentRow.Cells[8].Value.ToString());
            InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            WordDocument.Fields.Update();
            WordApplication.Visible = true;
        }
        public void ExportAutoToWordPage(DataGridView dgv)
        {
            Microsoft.Office.Interop.Word.Application WordApplication = new Microsoft.Office.Interop.Word.Application();
            object fileName = "C:\\Users\\Lenovo\\Documents\\autos.docx";
            Microsoft.Office.Interop.Word.Document WordDocument = WordApplication.Documents.Open(fileName);
            WordDocument.Variables["AutoName"].Value = dgv.CurrentRow.Cells[2].Value.ToString();
            WordDocument.Variables["AutoModel"].Value = dgv.CurrentRow.Cells[1].Value.ToString();
            WordDocument.Variables["AutoBody"].Value = dgv.CurrentRow.Cells[3].Value.ToString();
            WordDocument.Variables["AutoFuel"].Value = dgv.CurrentRow.Cells[4].Value.ToString();
            WordDocument.Variables["TransmissionType"].Value = dgv.CurrentRow.Cells[5].Value.ToString();
            WordDocument.Variables["AutoYear"].Value = dgv.CurrentRow.Cells[6].Value.ToString();
            WordDocument.Variables["DriveType"].Value = dgv.CurrentRow.Cells[7].Value.ToString();
            Microsoft.Office.Interop.Word.Range docRange = WordDocument.Range();
            Microsoft.Office.Interop.Word.Shape newShape = WordDocument.Shapes[1];
            newShape.Fill.UserPicture(dgv.CurrentRow.Cells[8].Value.ToString());
            InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            WordDocument.Fields.Update();
            WordApplication.Visible = true;
        }

        public void ExportPartFromSearchToWordPage(DataGridView dgv)
        {
            Microsoft.Office.Interop.Word.Application WordApplication = new Microsoft.Office.Interop.Word.Application();
            object fileName = "C:\\Users\\Lenovo\\Documents\\partsearch.docx";
            Microsoft.Office.Interop.Word.Document WordDocument = WordApplication.Documents.Open(fileName);
            WordDocument.Variables["PartName"].Value = dgv.CurrentRow.Cells[1].Value.ToString();
            WordDocument.Variables["PartCode"].Value = dgv.CurrentRow.Cells[2].Value.ToString();
            WordDocument.Variables["Category"].Value = dgv.CurrentRow.Cells[3].Value.ToString();
            WordDocument.Variables["AutoName"].Value = dgv.CurrentRow.Cells[4].Value.ToString();
            WordDocument.Variables["AutoModel"].Value = dgv.CurrentRow.Cells[5].Value.ToString();
            WordDocument.Variables["PartCost"].Value = dgv.CurrentRow.Cells[6].Value.ToString();
            WordDocument.Variables["PartAmount"].Value = dgv.CurrentRow.Cells[7].Value.ToString();
            Microsoft.Office.Interop.Word.Range docRange = WordDocument.Range();
            Microsoft.Office.Interop.Word.Shape newShape = WordDocument.Shapes[1];
            newShape.Fill.UserPicture(dgv.CurrentRow.Cells[8].Value.ToString());
            InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            WordDocument.Fields.Update();
            WordApplication.Visible = true;
        }

        public void ExportAutoFromSearchToWordPage(DataGridView dgv)
        {
            Microsoft.Office.Interop.Word.Application WordApplication = new Microsoft.Office.Interop.Word.Application();
            object fileName = "C:\\Users\\Lenovo\\Documents\\autos.docx";
            Microsoft.Office.Interop.Word.Document WordDocument = WordApplication.Documents.Open(fileName);
            WordDocument.Variables["AutoName"].Value = dgv.CurrentRow.Cells[4].Value.ToString();
            WordDocument.Variables["AutoModel"].Value = dgv.CurrentRow.Cells[1].Value.ToString();
            WordDocument.Variables["AutoBody"].Value = dgv.CurrentRow.Cells[3].Value.ToString();
            WordDocument.Variables["AutoFuel"].Value = dgv.CurrentRow.Cells[2].Value.ToString();
            WordDocument.Variables["TransmissionType"].Value = dgv.CurrentRow.Cells[5].Value.ToString();
            WordDocument.Variables["AutoYear"].Value = dgv.CurrentRow.Cells[7].Value.ToString();
            WordDocument.Variables["DriveType"].Value = dgv.CurrentRow.Cells[6].Value.ToString();
            Microsoft.Office.Interop.Word.Range docRange = WordDocument.Range();
            Microsoft.Office.Interop.Word.Shape newShape = WordDocument.Shapes[1];
            newShape.Fill.UserPicture(dgv.CurrentRow.Cells[8].Value.ToString());
            InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            WordDocument.Fields.Update();
            WordApplication.Visible = true;
        }

        private void CancelAddCategoryButton_Click(object sender, EventArgs e)
        {
            CancelAddInfo(CategoryAddPanel);
            CategoryAddPanel.Visible = false;
            OnButtons();
        }

        public bool CheckCategoryExists(string cName)
        {
            var command = new OdbcCommand($"select 1 from categories where category_name = '{cName}'", connection);
            return 1 == Convert.ToInt32(command.ExecuteScalar());
        }

        private void NewCategoryButton_Click(object sender, EventArgs e)
        {
            if (!CheckFields(CategoryAddPanel))
            {
                ThrowError("Поле имени категории для добавления пусто. Повторите ввод.");
                return;
            }
            if (CheckCategoryExists(CategoryAddTB.Text))
            {
                ThrowError("Такая категория уже существует. Повторите ввод.");
                return;
            }
            var insert = new OdbcCommand($"insert into categories values (default, '{CategoryAddTB.Text}', '{CategoryEngAddTB.Text}');", connection);
            insert.ExecuteNonQuery();
            ShowSuccess("Категория успешно добавлена.");
            this.categoriesTableAdapter.Fill(this.vvgcarpartsDataSet.categories);

        }

        private void CategoryAddPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(CategoryAddPanel);
        }

        private void CategoryAddPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                CategoryAddPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void DetailAddPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(DetailAddPanel);
        }

        private void DetailAddPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                DetailAddPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void ICFOkButton_Click(object sender, EventArgs e)
        {
            InternetConnectionFailedPanel.Visible = false;
        }

        private void FindInfoPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(FindInfoPanel);
        }

        private void FindInfoPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                FindInfoPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }

        private void TablesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MonthComboBox_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void ExportDate_Click(object sender, EventArgs e)
        {
        }

        private void MonthComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void AdvancedDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void PartCodeSellsViewCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            var command = new OdbcDataAdapter($"select distinct auto_name from autos inner join parts on autos.id = parts.car_id where part_name = '{PartNameSellsComboBox.Text}' and part_code = '{PartCodeSellsViewCB.Text}';", connection);
            var columns = new OdbcCommandBuilder(command);

            DataSet ds = new DataSet();

            command.Fill(ds, "autos");

            AutoNameSellsComboBox.DataSource = ds.Tables[0];
            AutoNameSellsComboBox.DisplayMember = "auto_name";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void PartPhotosDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ConfirmPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(ConfirmPanel);
        }

        private void DeleteAutoPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ChangeCoordinates(DeleteAutoPanel);
        }

        private void DeleteAutoPanel_MouseMove(object sender, MouseEventArgs e)
        {
            int iMouseX2 = MousePosition.X;
            int iMouseY2 = MousePosition.Y;
            if (e.Button == MouseButtons.Left)
                DeleteAutoPanel.Location = new System.Drawing.Point(iFormX + (iMouseX2 - iMouseX), iFormY + (iMouseY2 - iMouseY));
        }
    }
}
