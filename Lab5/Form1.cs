using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace Lab5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\Data\\Uni\\Year3\\Part1\\DataBases\\Lab7_5\\Lab7.accdb;Persist Security Info=False;";
        public OleDbConnection connection;
        public string query;
        public OleDbCommand command;
        public OleDbDataAdapter adapter;
        public DataTable dataTable;
        OleDbDataAdapter table1, table2, table3, table4, table5, table6, table7, table8, table9;



        DataTable table1DS, table2DS, table3DS, table4DS, table5DS, table6DS, table7DS, table8DS, table9DS;



        public void ConnectDB()
        {
            connection = new OleDbConnection(connectString);
            connection.Open();
            SetTables();
            table1DS = new DataTable();
            table2DS = new DataTable();
            table3DS = new DataTable();
            table4DS = new DataTable();
            table5DS = new DataTable();
            table6DS = new DataTable();
            table7DS = new DataTable();
            table8DS = new DataTable();
            table9DS = new DataTable();
            table1.Fill(table1DS);
            table2.Fill(table2DS);
            table3.Fill(table3DS);
            table4.Fill(table4DS);
            table5.Fill(table5DS);
            table6.Fill(table6DS);
            table7.Fill(table7DS);
            table8.Fill(table8DS);
            table9.Fill(table9DS);
            table2DS.Columns[0].AutoIncrement = true;
            table3DS.Columns[0].AutoIncrement = true;
            table4DS.Columns[0].AutoIncrement = true;
            table5DS.Columns[0].AutoIncrement = true;
            table6DS.Columns[0].AutoIncrement = true;
            table9DS.Columns[0].AutoIncrement = true;

            dataGridView1.DataSource = table1DS;
            dataGridView2.DataSource = table2DS;
            dataGridView3.DataSource = table3DS;
            dataGridView4.DataSource = table4DS;
            dataGridView5.DataSource = table5DS;
            dataGridView6.DataSource = table6DS;
            dataGridView7.DataSource = table7DS;
            dataGridView8.DataSource = table8DS;
            dataGridView9.DataSource = table9DS;
        }
        private void SetTables()
        {
            string sqlTable1 = "SELECT * FROM Автор";
            string sqlTable2 = "SELECT * FROM Бібліотекар";
            string sqlTable3 = "SELECT * FROM Заказ";
            string sqlTable4 = "SELECT * FROM Зберігання";
            string sqlTable5 = "SELECT * FROM Книга";
            string sqlTable6 = "SELECT * FROM Працівник";
            string sqlTable7 = "SELECT * FROM Напрямок";
            string sqlTable8 = "SELECT * FROM Пише";
            string sqlTable9 = "SELECT * FROM Сховище";

            table1 = new OleDbDataAdapter(sqlTable1, connection);
            table2 = new OleDbDataAdapter(sqlTable2, connection);
            table3 = new OleDbDataAdapter(sqlTable3, connection);
            table4 = new OleDbDataAdapter(sqlTable4, connection);
            table5 = new OleDbDataAdapter(sqlTable5, connection);
            table6 = new OleDbDataAdapter(sqlTable6, connection);
            table7 = new OleDbDataAdapter(sqlTable7, connection);
            table8 = new OleDbDataAdapter(sqlTable8, connection);
            table9 = new OleDbDataAdapter(sqlTable9, connection);
        }
        public void CloseDB()
        {
            connection.Close();
        }

        public void RefreshTables()
        {
            table1DS.Rows.Clear();
            table1.Fill(table1DS);
            table2DS.Rows.Clear();
            table2.Fill(table2DS);
            table3DS.Rows.Clear();
            table3.Fill(table3DS);
            table4DS.Rows.Clear();
            table4.Fill(table4DS);
            table5DS.Rows.Clear();
            table5.Fill(table5DS);
            table6DS.Rows.Clear();
            table6.Fill(table6DS);
            table7DS.Rows.Clear();
            table7.Fill(table7DS);
            table8DS.Rows.Clear();
            table8.Fill(table8DS);
            table9DS.Rows.Clear();
            table9.Fill(table9DS);


        }



        private void Form1_Load(object sender, EventArgs e)
        {
            tabControl1.Height = this.Height - 100;
            ConnectDB();
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            switch (index)
            {
                case 9:
                    string input = Interaction.InputBox("Введіть Напрям", "Напрям", "Novella");
                    query = "SELECT Пише.[Назва Книги], Пише.[Назва Напряму] FROM Пише WHERE (((Пише.[Назва Напряму])='" + input + "'));";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView10.DataSource = dataTable;
                    break;
                case 10:
                    query = "SELECT Бібліотекар.[ID Бібліотекара], Бібліотекар.Прізвище, Count(Заказ.[ID Бібліотекара]) AS [Кількість Заказів]FROM Книга INNER JOIN (Бібліотекар INNER JOIN Заказ ON Бібліотекар.[ID Бібліотекара] = Заказ.[ID Бібліотекара]) ON Книга.[ID Книги] = Заказ.[ID Кніги]GROUP BY Бібліотекар.[ID Бібліотекара], Бібліотекар.Прізвище;";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView11.DataSource = dataTable;
                    break;
                case 11:
                    query = "SELECT First(Автор.[Ім'я Автора]) AS [Ім'я Автора], First(Кількість) AS [Разів Обрано] FROM (SELECT Count(Працівник.[Стаж Роботи]) AS Кількість, Автор.[Ім'я Автора] FROM (Автор INNER JOIN Пише ON Автор.[Ім'я Автора] = Пише.[Ім'я Автора]) INNER JOIN (Книга INNER JOIN (Працівник INNER JOIN Заказ ON Працівник.[ID працівника] = Заказ.[ID Працівника]) ON Книга.[ID Книги] = Заказ.[ID Кніги]) ON Пише.[ISBN Книги] = Книга.ISBN WHERE (((Працівник.[Стаж Роботи])>(SELECT Avg([Стаж Роботи]) FROM Працівник))) GROUP BY Автор.[Ім'я Автора] ORDER BY Count(Працівник.[Стаж Роботи]) DESC)  AS [%$##@_Alias];";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView12.DataSource = dataTable;
                    break;

            }

        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            tabControl1.Height = this.Height - 100;
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            //Deletes Data Base entry
            int testint = 0;
            bool testconvert;
            int TabIndex = tabControl1.SelectedIndex;
            string indexName = "";
            string indexValue = "";
            string tableName = tabControl1.SelectedTab.Text;
            switch (TabIndex)
            {
                case 0:
                    indexValue = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView1.Columns[0].HeaderText;
                    break;
                case 1:
                    indexValue = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView2.Columns[0].HeaderText;
                    break;
                case 2:
                    indexValue = dataGridView3.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView3.Columns[0].HeaderText;
                    break;
                case 3:
                    indexValue = dataGridView4.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView4.Columns[0].HeaderText;
                    break;
                case 4:
                    indexValue = dataGridView5.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView5.Columns[0].HeaderText;
                    break;
                case 5:
                    indexValue = dataGridView6.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView6.Columns[0].HeaderText;
                    break;
                case 6:
                    indexValue = dataGridView7.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView7.Columns[0].HeaderText;
                    break;
                case 7:
                    indexValue = dataGridView8.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView8.Columns[0].HeaderText;
                    break;
                case 8:
                    indexValue = dataGridView9.SelectedRows[0].Cells[0].Value.ToString();
                    indexName = dataGridView9.Columns[0].HeaderText;
                    break;

            }
            testconvert = int.TryParse(indexValue, out testint);
            if (testconvert == true)
            {
                query = "DELETE FROM " + tableName + " WHERE [" + indexName + "]=" + indexValue + ";";
            }
            else
            {
                query = "DELETE FROM " + tableName + " WHERE [" + indexName + "]='" + indexValue + "';";
            }

            OleDbCommand command = new OleDbCommand(query, connection);
            command.ExecuteNonQuery();
            RefreshTables();

        }


        private void EditButton_Click(object sender, EventArgs e)
        {
            //Edits Data Base entry
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            string indexValue = "";
            string indexName = "";
            bool test = false;
            List<TextBox> texts = new List<TextBox>();
            List<Label> labels = new List<Label>();
            List<String> columns = new List<String>();
            //Building Prompt
            Form prompt = new Form();
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            prompt.MaximizeBox = false;
            prompt.MinimizeBox = false;
            prompt.Width = 500;
            prompt.Height = 500;
            Button confirm = new Button() { Text = "Confirm", Left = prompt.Width / 2 - 150, Width = 100, Top = prompt.Height - 130, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = prompt.Width / 2 + 50, Width = 100, Top = prompt.Height - 130, Height = 50, DialogResult = DialogResult.Cancel };
            prompt.Controls.Add(confirm);
            prompt.Controls.Add(cancel);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;
            //Checking which table

            try
            {

                switch (tabIndex)

                {
                    case 0:

                        for (int i = 1; i < dataGridView1.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView1.Columns[0].HeaderText.ToString();
                        break;
                    case 1:
                        for (int i = 1; i < dataGridView2.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView2.Columns[0].HeaderText.ToString();
                        break;
                    case 2:
                        for (int i = 1; i < dataGridView3.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView3.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView3.Columns[0].HeaderText.ToString();
                        break;
                    case 3:
                        for (int i = 1; i < dataGridView4.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView4.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView4.Columns[0].HeaderText.ToString();
                        break;
                    case 4:
                        for (int i = 1; i < dataGridView5.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView5.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView5.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView5.Columns[0].HeaderText.ToString();
                        break;
                    case 5:
                        for (int i = 1; i < dataGridView6.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView6.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView6.Columns[0].HeaderText.ToString();
                        break;
                    case 6:
                        for (int i = 1; i < dataGridView7.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView7.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 250, Height = 100, Top = 70 + (30 * i), Text = dataGridView7.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView7.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView7.Columns[0].HeaderText.ToString();
                        break;
                    case 7:
                        for (int i = 1; i < dataGridView8.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView8.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView8.Columns[0].HeaderText.ToString();
                        break;
                    case 8:
                        for (int i = 1; i < dataGridView9.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.Columns[i].HeaderText.ToString() });
                            texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.SelectedRows[0].Cells[i].Value.ToString() });
                            columns.Add(dataGridView9.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView9.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView9.Columns[0].HeaderText.ToString();
                        break;

                }
                foreach (TextBox text in texts)
                {
                    prompt.Controls.Add(text);
                }
                foreach (Label label in labels)
                {
                    prompt.Controls.Add(label);
                }
                prompt.ShowDialog();
                if (prompt.DialogResult == DialogResult.OK)
                {
                    test = int.TryParse(indexValue, out _);
                    query = "UPDATE [" + tabControl1.SelectedTab.Text + "] SET ";
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        if (test == true)
                        {
                            query += "[" + columns[i] + "] = '" + texts[i].Text + "'";
                        }
                        else
                        {
                            query += "[" + columns[i] + "] = '" + texts[i].Text + "'";
                        }
                        if (i != columns.Count() - 1)
                        {
                            query += ", ";
                        }
                    }
                    if (test == true)
                    {
                        query += " WHERE [" + indexName + "] = " + indexValue + ";";
                    }
                    else
                    {
                        query += " WHERE [" + indexName + "] = '" + indexValue + "';";
                    }


                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.ExecuteNonQuery();
                    RefreshTables();
                    prompt.Close();
                    return;
                }
                if (prompt.DialogResult == DialogResult.Cancel)
                {
                    prompt.Close();

                }


            }
            catch
            {
                MessageBox.Show("Error: You are trying to edit while you didn't choose a row!");

            }

        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            string indexValue = "", indexName = "";
            List<TextBox> texts = new List<TextBox>();
            List<Label> labels = new List<Label>();
            List<String> columns = new List<String>();
            int j = 0;
            if (tabIndex == 0 || tabIndex == 6 || tabIndex == 7)
            {
                j = 1;
            }
            //Building Prompt
            Form promptAdd = new Form();
            promptAdd.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptAdd.StartPosition = FormStartPosition.CenterScreen;
            promptAdd.MaximizeBox = false;
            promptAdd.MinimizeBox = false;
            promptAdd.Width = 500;
            promptAdd.Height = 500;
            Button confirm = new Button() { Text = "Confirm", Left = promptAdd.Width / 2 - 150, Width = 100, Top = promptAdd.Height - 130, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = promptAdd.Width / 2 + 50, Width = 100, Top = promptAdd.Height - 130, Height = 50, DialogResult = DialogResult.Cancel };
            promptAdd.Controls.Add(confirm);
            promptAdd.Controls.Add(cancel);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;

            switch (tabIndex)

            {
                case 0:

                    for (int i = 1 - j; i < dataGridView1.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                    }
                    break;
                case 1:
                    for (int i = 1 - j; i < dataGridView2.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                    }
                    break;
                case 2:
                    for (int i = 1 - j; i < dataGridView3.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                    }

                    break;
                case 3:
                    for (int i = 1 - j; i < dataGridView4.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                    }

                    break;
                case 4:
                    for (int i = 1 - j; i < dataGridView5.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView5.Columns[i].HeaderText.ToString());

                    }

                    break;
                case 5:
                    for (int i = 1 - j; i < dataGridView6.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                    }

                    break;
                case 6:
                    for (int i = 1 - j; i < dataGridView7.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView7.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 250, Height = 100, Top = 70 + (30 * i) });
                        columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                    }

                    break;
                case 7:
                    for (int i = 1 - j; i < dataGridView8.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                    }

                    break;
                case 8:
                    for (int i = 1 - j; i < dataGridView9.ColumnCount; i++)
                    {
                        labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.Columns[i].HeaderText.ToString() });
                        texts.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                        columns.Add(dataGridView9.Columns[i].HeaderText.ToString());

                    }

                    break;

            }
            foreach (TextBox text in texts)
            {
                promptAdd.Controls.Add(text);
            }
            foreach (Label label in labels)
            {
                promptAdd.Controls.Add(label);
            }
            promptAdd.ShowDialog();

            if (promptAdd.DialogResult == DialogResult.OK)
            {
                query = "INSERT INTO [" + tabControl1.SelectedTab.Text + "] (";
                for (int i = 0; i < columns.Count(); i++)
                {
                    query += "[" + columns[i] + "]";

                    if (i != columns.Count() - 1)
                    {
                        query += ", ";
                    }
                }
                query += ")  VALUES (";
                for (int i = 0; i < columns.Count(); i++)
                {
                    query += "'" + texts[i].Text + "'";
                    if (i != columns.Count() - 1)
                    {
                        query += ", ";
                    }
                }
                query += ");";

                OleDbCommand command = new OleDbCommand(query, connection);
                command.ExecuteNonQuery();
                RefreshTables();
                promptAdd.Close();
                return;
            }
            if (promptAdd.DialogResult == DialogResult.Cancel)
            {
                promptAdd.Close();

            }
        }

        private Form gridForm(DataGridView searchView)
        {

            //New Form for showing results of search
            Form SearchResult = new Form();
            SearchResult.FormBorderStyle = FormBorderStyle.FixedDialog;
            SearchResult.StartPosition = FormStartPosition.CenterScreen;
            SearchResult.Width = 750;
            SearchResult.Height = 500;
            searchView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            searchView.Dock = DockStyle.Fill;
            searchView.ReadOnly = true;
            searchView.AllowUserToAddRows = false;
            searchView.AllowUserToDeleteRows = false;
            SearchResult.Controls.Add(searchView);
            return SearchResult;
        }


        private void searchButton_Click(object sender, EventArgs e)
        {
            //First check if anything is written in the search field
            if (searchBox.Text == "")
            {
                MessageBox.Show("Error, search field cannot be empty!");
                return;
            }

            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            List<String> columns = new List<String>();
            bool isString = true, isInt = false, isDate = false;
            string text = searchBox.Text;
            //New Table for Searching
            OleDbDataAdapter tableSearch;
            DataTable tableSearchData = new DataTable();
            DataGridView searchView = new DataGridView();
            if (int.TryParse(text, out _))
            {
                isInt = true;
                isString = false;
                isDate = false;
            }
            else if (System.DateTime.TryParse(text, out _))
            {
                isInt = false;
                isString = false;
                isDate = true;
            }

            switch (tabIndex)

            {
                case 0:

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView1.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView1.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                        }
                    }
                    break;
                case 1:
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView2.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView2.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                        }
                    }
                    break;
                case 2:
                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                    {
                        if (dataGridView3.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView3.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView3.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 3:
                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                    {
                        if (dataGridView4.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView4.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView4.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 4:
                    for (int i = 0; i < dataGridView5.ColumnCount; i++)
                    {
                        if (dataGridView5.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView5.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView5.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView5.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 5:
                    for (int i = 0; i < dataGridView6.ColumnCount; i++)
                    {
                        if (dataGridView6.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView6.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView6.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 6:
                    for (int i = 0; i < dataGridView7.ColumnCount; i++)
                    {
                        if (dataGridView7.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView7.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView7.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 7:
                    for (int i = 0; i < dataGridView8.ColumnCount; i++)
                    {
                        if (dataGridView8.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView8.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView8.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 8:
                    for (int i = 0; i < dataGridView9.ColumnCount; i++)
                    {
                        if (dataGridView9.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView9.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView9.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView9.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
            }
            query = "SELECT * FROM [" + tabControl1.SelectedTab.Text.ToString() + "] WHERE ";
            for (int i = 0; i < columns.Count; i++)
            {
                if (isInt == true)
                {
                    query += "[" + columns[i] + "] = " + searchBox.Text + "";
                }
                else if (isDate == true)
                {
                    query += "[" + columns[i] + "] = #" + searchBox.Text + "#";
                }
                else
                {
                    query += "[" + columns[i] + "] = '" + searchBox.Text + "'";
                }

                if (i != columns.Count - 1)
                {
                    query += " OR ";
                }
            }
            query += ";";

            searchView.DataSource = null;
            tableSearch = new OleDbDataAdapter(query, connection);
            tableSearch.Fill(tableSearchData);
            searchView.DataSource = tableSearchData;
            Form SearchResult = gridForm(searchView);
            SearchResult.Show();
            RefreshTables();
        }
        static bool isAscending = true;
        private void sortButton_Click(object sender, EventArgs e)
        {
            //Variables

            int tabIndex = tabControl1.SelectedIndex;
            string column = "";

            switch (tabIndex)
            {
                case 0:
                    column = dataGridView1.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 1:
                    column = dataGridView2.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 2:
                    column = dataGridView3.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 3:
                    column = dataGridView4.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 4:
                    column = dataGridView5.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 5:
                    column = dataGridView6.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 6:
                    column = dataGridView7.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 7:
                    column = dataGridView8.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 8:
                    column = dataGridView9.CurrentCell.OwningColumn.HeaderText;
                    break;

            }
            if (isAscending == true)
            {
                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text + "] ORDER BY [" + column + "] ASC;";
                isAscending = false;
            }
            else
            {
                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text + "] ORDER BY [" + column + "] DESC;";
                isAscending = true;
            }
            switch (tabIndex)

            {
                case 0:
                    table1 = new OleDbDataAdapter(query, connection);
                    break;
                case 1:
                    table2 = new OleDbDataAdapter(query, connection);
                    break;
                case 2:
                    table3 = new OleDbDataAdapter(query, connection);
                    break;
                case 3:
                    table4 = new OleDbDataAdapter(query, connection);
                    break;
                case 4:
                    table5 = new OleDbDataAdapter(query, connection);
                    break;
                case 5:
                    table6 = new OleDbDataAdapter(query, connection);
                    break;
                case 6:
                    table7 = new OleDbDataAdapter(query, connection);
                    break;
                case 7:
                    table8 = new OleDbDataAdapter(query, connection);
                    break;
                case 8:
                    table9 = new OleDbDataAdapter(query, connection);
                    break;

            }

            RefreshTables();

        }

        private void filterButton_Click(object sender, EventArgs e)
        {

            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            Type ColumnValueType = null;
            string column = "";
            string FilterValue;
            bool isString = true, isInt = false, isDate = false;
            string filterSign = "";
            //Prompt Form
            Form promptFilter = new Form();
            promptFilter.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptFilter.StartPosition = FormStartPosition.CenterScreen;
            promptFilter.MaximizeBox = false;
            promptFilter.MinimizeBox = false;
            promptFilter.Width = 500;
            promptFilter.Height = 250;
            promptFilter.Text = "Filter";
            RadioButton BiggerButton = new RadioButton() { Text = "Bigger", Left = promptFilter.Width / 2 - 175, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            RadioButton SmallerButton = new RadioButton() { Text = "Smaller", Left = promptFilter.Width / 2 - 50, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            RadioButton EqualsButton = new RadioButton() { Text = "Equals", Left = promptFilter.Width / 2 + 75, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            Button confirm = new Button() { Text = "Confirm", Left = promptFilter.Width / 2 - 150, Width = 100, Top = promptFilter.Height - 100, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = promptFilter.Width / 2 + 50, Width = 100, Top = promptFilter.Height - 100, Height = 50, DialogResult = DialogResult.Cancel };
            Label promptLabel = new Label() { Text = "Input Value", Left = promptFilter.Width / 2 - 150, Top = promptFilter.Height - 175, BackColor = System.Drawing.Color.Transparent };
            TextBox promptText = new TextBox() { Left = promptFilter.Width / 2 - 150, Top = promptFilter.Height - 150, Width = promptFilter.Width / 2 };
            promptFilter.Controls.Add(confirm);
            promptFilter.Controls.Add(cancel);
            promptFilter.Controls.Add(promptLabel);
            promptFilter.Controls.Add(promptText);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;

            switch (tabIndex)
            {
                case 0:
                    column = dataGridView1.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView1.CurrentCell.OwningColumn.ValueType;
                    break;
                case 1:
                    column = dataGridView2.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView2.CurrentCell.OwningColumn.ValueType;
                    break;
                case 2:
                    column = dataGridView3.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView3.CurrentCell.OwningColumn.ValueType;
                    break;
                case 3:
                    column = dataGridView4.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView4.CurrentCell.OwningColumn.ValueType;
                    break;
                case 4:
                    column = dataGridView5.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView5.CurrentCell.OwningColumn.ValueType;
                    break;
                case 5:
                    column = dataGridView6.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView6.CurrentCell.OwningColumn.ValueType;
                    break;
                case 6:
                    column = dataGridView7.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView7.CurrentCell.OwningColumn.ValueType;
                    break;
                case 7:
                    column = dataGridView8.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView8.CurrentCell.OwningColumn.ValueType;
                    break;
                case 8:
                    column = dataGridView9.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView9.CurrentCell.OwningColumn.ValueType;
                    break;

            }

            if(ColumnValueType != Type.GetType("System.String"))
            {
                promptFilter.Controls.Add(BiggerButton);
                promptFilter.Controls.Add(SmallerButton);
                promptFilter.Controls.Add(EqualsButton);

            }


            promptFilter.ShowDialog();
            if (promptFilter.DialogResult != DialogResult.OK)
            {
                promptFilter.Close();
                return;
            }
            FilterValue = promptText.Text;
            if (ColumnValueType != Type.GetType("System.String"))
            {
                if (BiggerButton.Checked)
                {
                    filterSign = ">";

                }
                if (SmallerButton.Checked)
                {
                    filterSign = "<";

                }
                if (EqualsButton.Checked)
                {
                    filterSign = "=";
                }

                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text.ToString() + "] WHERE [" + column + "] " + filterSign + " ";
                if (ColumnValueType == Type.GetType("System.Int32"))
                {
                    query += FilterValue;
                }
                else if (ColumnValueType == Type.GetType("System.DateTime"))
                {
                    query += "#" + FilterValue + "#";
                }
                else
                {
                    query += "'" + FilterValue + "'";
                }
                query += ";";
            }
            else
            {
                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text.ToString() + "] WHERE [" + column + "] LIKE '%" + FilterValue + "%';";
            }
           
            switch (tabIndex)
            {
                case 0:
                    table1 = new OleDbDataAdapter(query, connection);
                    break;
                case 1:
                    table2 = new OleDbDataAdapter(query, connection);
                    break;
                case 2:
                    table3 = new OleDbDataAdapter(query, connection);
                    break;
                case 3:
                    table4 = new OleDbDataAdapter(query, connection);
                    break;
                case 4:
                    table5 = new OleDbDataAdapter(query, connection);
                    break;
                case 5:
                    table6 = new OleDbDataAdapter(query, connection);
                    break;
                case 6:
                    table7 = new OleDbDataAdapter(query, connection);
                    break;
                case 7:
                    table8 = new OleDbDataAdapter(query, connection);
                    break;
                case 8:
                    table9 = new OleDbDataAdapter(query, connection);
                    break;

            }

            RefreshTables();

        }
        private void resetButton_Click(object sender, EventArgs e)
        {
            SetTables();
            RefreshTables();
        }
    }
}

