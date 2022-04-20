using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompOut
{
    public partial class Form_proizv : Form
    {
        public Form_proizv()
        {
            InitializeComponent();
        }
        DataSet ds;
        SqlDataAdapter adapter;
        private void Form_proizv_Load(object sender, EventArgs e)
        {
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;

            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                adapter = new SqlDataAdapter(SQL_query.sql_manuf, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                dataGridView1.RowHeadersVisible = false;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[1].HeaderText = "Наименование производителя";
                dataGridView1.Columns[2].HeaderText = "Страна производства";
                dataGridView1.Columns[1].Width = dataGridView1.Width / 2;
                dataGridView1.Columns[2].Width = dataGridView1.Width / 2;

            }
        }
        // добавить производителя
        private void button16_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox85.Text.Equals(""))
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.New_manuf, textBox1.Text, textBox85.Text);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.sql_manuf);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    textBox1.Text = "";
                    textBox85.Text = "";
                }

            }
        }
        //удалить производителя
        private void button15_Click(object sender, EventArgs e)
        {
            int index = dataGridView1.CurrentCell.RowIndex; // номер выделенной строчки
            if (index >= 0)
            {
                string id = dataGridView1.Rows[index].Cells[0].Value.ToString(); // получаем id выделенной строчки
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.Delete_manuf, id);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.sql_manuf);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }
        }
    }
}
