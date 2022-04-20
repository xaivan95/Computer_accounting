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
    public partial class Form_compon : Form
    {
        public Form_compon()
        {
            InitializeComponent();
        }
        List<Data> lists = new List<Data>();
        List<TabPage> tabs = new List<TabPage>();
        DataSet[] ds = new DataSet[24];
        SqlDataAdapter adap_button = new SqlDataAdapter();
        DataSet ds_button = new DataSet();
        SqlDataAdapter[] adapter = new SqlDataAdapter[24];
        List<string> Tables = new List<string>() { "cpu", "mb", "memory", "video", "sound", "cases", "hdd", "fdd", "lan", "cdr", "cdrw", "dvd", "display", "printer", "scaner", "ups", "mouse", "key", "usb", "modem", "other", "soft" };
        List<string> Item = new List<string>() { "cpu", "mb", "memory", "video", "sound", "case", "hdd", "fdd", "lan", "cdr", "cdrw", "dvd", "display", "printer", "scaner", "ups", "mouse", "key", "usb", "modem", "other", "soft" };
        int ID_select;
        Dictionary<string, string[]> compon_comb = new Dictionary<string, string[]>()
        {
            {"dataGridView1", new string[] {"textBox1", "textBox2", "textBox3", "comboBox1" } },
            {"dataGridView2", new string[] {"textBox6", "textBox4", "comboBox2" } },
            {"dataGridView3", new string[] {"textBox9", "textBox8", "textBox5", "textBox7", "comboBox3" } },
            {"dataGridView4", new string[] {"textBox12", "textBox11", "textBox10", "comboBox4" } },
            {"dataGridView5", new string[] {"textBox15", "textBox13", "comboBox5" } },
            {"dataGridView6", new string[] {"textBox18", "textBox17", "textBox16", "comboBox6" } },
            {"dataGridView7", new string[] {"textBox21", "textBox20", "textBox19", "comboBox7" } },
            {"dataGridView8", new string[] {"textBox24", "textBox23", "textBox22", "comboBox8" } },
            {"dataGridView9", new string[] {"textBox27", "textBox26", "textBox25", "comboBox9" } },

            {"dataGridView10", new string[] {"textBox30", "comboBox23", "textBox28", "comboBox10" } },
            {"dataGridView11", new string[] { "textBox33", "comboBox24", "comboBox25", "textBox31", "comboBox11" } },
            {"dataGridView12", new string[] { "textBox36", "comboBox27", "comboBox26", "comboBox29", "comboBox28", "textBox34", "comboBox12" } },

            {"dataGridView13", new string[] {"textBox39", "textBox14", "textBox38", "textBox37", "comboBox13" } },
            {"dataGridView14", new string[] {"textBox42", "comboBox30", "textBox41", "textBox29", "textBox40", "comboBox14" } },
            {"dataGridView15", new string[] { "textBox45", "comboBox31", "textBox32", "textBox43", "comboBox15" } },
            {"dataGridView16", new string[] { "textBox48", "textBox47", "textBox46", "comboBox16" } },
            {"dataGridView17", new string[] {"textBox51", "textBox49",  "comboBox17" } },
            {"dataGridView18", new string[] {"textBox54", "textBox52",  "comboBox18" } },
            {"dataGridView19", new string[] {"textBox57", "textBox56", "textBox55", "comboBox19" } },
            {"dataGridView20", new string[] {"textBox60", "textBox59", "textBox58", "comboBox20" } },
            {"dataGridView21", new string[] { "textBox63", "textBox61","comboBox21" } },

            {"dataGridView22", new string[] { "textBox66", "textBox64",  "comboBox22" } }
        };


        private void Form_compon_Load(object sender, EventArgs e)
        {
            Control[] contrl;
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {

                connection.Open();
                for (int i = 1; i < 23; i++)
                {
                    ds[i - 1] = new DataSet();

                    contrl = Controls.Find("dataGridView" + i.ToString(), true);
                    (contrl[0] as DataGridView).SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    (contrl[0] as DataGridView).AllowUserToAddRows = false;

                    adapter[i - 1] = new SqlDataAdapter(String.Format(SQL_query.sql_tables_components, Tables[i - 1], Item[i - 1]), connection);
                    adapter[i - 1].Fill(ds[i - 1]);
                    (contrl[0] as DataGridView).DataSource = ds[i - 1].Tables[0];
                    (contrl[0] as DataGridView).RowHeadersVisible = false;
                    (contrl[0] as DataGridView).Columns[0].Visible = false;
                    (contrl[0] as DataGridView).Tag = i;
                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_manuf_id"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_manuf_id"].Visible = false;
                    else
                        (contrl[0] as DataGridView).Columns["manuf_id"].Visible = false;

                    if ((contrl[0] as DataGridView).Columns.Contains("manuf_name"))
                        (contrl[0] as DataGridView).Columns["manuf_name"].HeaderText = "Производитель";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_model"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_model"].HeaderText = "Модель";
                    else
                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_name"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_name"].HeaderText = "Название";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_note"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_note"].HeaderText = "Характеристика";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_speed"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_speed"].HeaderText = "Скорость";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_size"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_size"].HeaderText = "Объем памяти";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_type"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_type"].HeaderText = "Тип";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_resolution"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_resolution"].HeaderText = "Разрешение";

                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_read"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_read"].HeaderText = "Скорость чтения";
                    if ((contrl[0] as DataGridView).Columns.Contains(Item[i - 1] + "_write"))
                        (contrl[0] as DataGridView).Columns[Item[i - 1] + "_write"].HeaderText = "Скорость записи";
                }

                ds[23] = new DataSet();
                adapter[23] = new SqlDataAdapter(SQL_query.sql_manuf, connection);
                adapter[23].Fill(ds[23]);
            }
            if (dataGridView13.Columns.Contains("display_size"))
                dataGridView13.Columns["display_size"].HeaderText = "Размер экрана";

            if (dataGridView16.Columns.Contains("ups_power"))
                dataGridView16.Columns["ups_power"].HeaderText = "Мощность";

            if (dataGridView12.Columns.Contains("CD_read"))
                dataGridView12.Columns["CD_read"].HeaderText = "Скорость чтения CD";
            if (dataGridView12.Columns.Contains("CD_write"))
                dataGridView12.Columns["CD_write"].HeaderText = "Скорость записи CD";
            if (dataGridView12.Columns.Contains("DVD_read"))
                dataGridView12.Columns["DVD_read"].HeaderText = "Скорость чтения DVD";
            if (dataGridView12.Columns.Contains("DVD_write"))
                dataGridView12.Columns["DVD_write"].HeaderText = "Скорость записи DVD";



            foreach (DataRow row in ds[23].Tables[0].Rows)
            {
                lists.Add(new Data(row.Field<int>("manuf_id"), row.Field<string>("manuf_name")));
            }
            for (int i = 0; i < 22; i++)
            {
                contrl = Controls.Find("comboBox" + (i + 1).ToString(), true);
                (contrl[0] as ComboBox).DataSource = lists;
                (contrl[0] as ComboBox).DisplayMember = "Name";
                (contrl[0] as ComboBox).ValueMember = "id";
            }



            tabs.Add(tabPage1);
            for (int i = 1; i < 22; i++)
            {
                tabs.Add(tabControl1.TabPages[1]);
                HidePage(tabControl1.TabPages[1]);
            }


        }
        public void HidePage(TabPage tabpage4)
        {
            if (!tabControl1.TabPages.Contains(tabpage4)) return;
            tabControl1.TabPages.Remove(tabpage4);
        }

        public void ShowPage(TabPage tabpage4)
        {
            if (tabControl1.TabPages.Contains(tabpage4)) return;
            tabControl1.TabPages.Add(tabpage4);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

            if (e.Node.Level == 1)
            {
                HidePage(tabControl1.TabPages[0]);
                for (int i = 0; i < 22; i++)
                {
                    if (tabs[i].Text.Equals(e.Node.Text))
                    {
                        ShowPage(tabs[i]);
                    }
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            Control[] contrl;
            if (((DataGridView)sender).SelectedCells.Count > 0)
            {
                DataGridViewSelectedCellCollection DGVCell = ((DataGridView)sender).SelectedCells;
                var dgvc = DGVCell[1];
                var dgvr = dgvc.OwningRow;
                byte i = 1;
                foreach (string s in compon_comb[((DataGridView)sender).Name])
                {
                    contrl = Controls.Find(s, true);
                    if (contrl.Count() > 0)
                        contrl[0].Text = dgvr.Cells[i].Value.ToString();
                    i++;
                }
                contrl = Controls.Find(compon_comb[((DataGridView)sender).Name][compon_comb[((DataGridView)sender).Name].Count() - 2], true);
                if (contrl.Count() > 0)
                    contrl[0].Text = dgvr.Cells[dgvr.Cells.Count - 2].Value.ToString();

                contrl = Controls.Find(compon_comb[((DataGridView)sender).Name][compon_comb[((DataGridView)sender).Name].Count() - 1], true);
                if (contrl.Count() > 0)
                    contrl[0].Text = dgvr.Cells[dgvr.Cells.Count - 1].Value.ToString();
                ID_select = (int) dgvr.Cells[0].Value;
            }
        }
        //удалить запись
        private void button3_Click(object sender, EventArgs e)
        {
            int b = int.Parse((string)((Button)sender).Tag);
            Control[] contrl;
            contrl = Controls.Find("dataGridView" + (b + 1).ToString(), true);
            var cell = (contrl[0] as DataGridView).SelectedCells;
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {

                connection.Open();
                var query = string.Format(SQL_query.Delete_complect, Tables[b], Item[b], cell[0].Value.ToString());
                adap_button = new SqlDataAdapter(query, connection);
                ds_button = new DataSet();
                adap_button.Fill(ds_button);

                adapter[b] = new SqlDataAdapter(String.Format(SQL_query.sql_tables_components, Tables[b], Item[b]), connection);
                ds[b] = new DataSet();
                adapter[b].Fill(ds[b]);
                (contrl[0] as DataGridView).DataSource = ds[b].Tables[0];

            }
        }
        //Добавить комплектующею
        private void button1_Click(object sender, EventArgs e)
        {
            Control[] contrl;
            int b = int.Parse((string)((Button)sender).Tag);
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                adapter[b] = new SqlDataAdapter(String.Format(SQL_query.sql_tables, Tables[b]), connection);
                ds[b]= new DataSet();
                adapter[b].Fill(ds[b]);           
                DataTable dt = ds[b].Tables[0];
                DataRow newRow = dt.NewRow();
                for (int i = 1; i <newRow.Table.Columns.Count-2; i++)
                {
                    contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][i-1], true);
                    if (contrl[0].GetType() == typeof(TextBox))
                    {
                        newRow[i] = contrl[0].Text;
                    }
                    else
                        if (contrl[0].GetType() == typeof(ComboBox))
                    {
                        newRow[i] = (contrl[0] as ComboBox).SelectedItem;
                    }
                }
                contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][compon_comb["dataGridView" + (b + 1).ToString()].Count()-1], true);
                newRow[newRow.Table.Columns.Count - 2] = (contrl[0] as ComboBox).SelectedValue;
                contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][compon_comb["dataGridView" + (b + 1).ToString()].Count() - 2], true);
                newRow[newRow.Table.Columns.Count - 1] = contrl[0].Text;
                dt.Rows.Add(newRow);
                SqlCommandBuilder command = new SqlCommandBuilder(adapter[b]);
                adapter[b].Update(ds[b]);

                adapter[b] = new SqlDataAdapter(String.Format(SQL_query.sql_tables_components, Tables[b], Item[b]), connection);
                ds[b] = new DataSet();
                adapter[b].Fill(ds[b]);
                contrl = Controls.Find("dataGridView" + (b + 1).ToString(), true);
                (contrl[0] as DataGridView).DataSource = ds[b].Tables[0];

            }
        }
        //Изменить
        private void button47_Click(object sender, EventArgs e)
        {
            Control[] contrl;
            int b = int.Parse((string)((Button)sender).Tag);
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                adapter[b] = new SqlDataAdapter(String.Format(SQL_query.sql_tables, Tables[b]), connection);
                ds[b] = new DataSet();
                adapter[b].Fill(ds[b]);
                DataTable dt = ds[b].Tables[0];
                DataRow newRow = dt.Select(Item[b]+"_id = " + ID_select.ToString())[0];
                for (int i = 1; i < newRow.Table.Columns.Count - 2; i++)
                {
                    contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][i - 1], true);
                    if (contrl[0].GetType() == typeof(TextBox))
                    {
                        newRow[i] = contrl[0].Text;
                    }
                    else
                        if (contrl[0].GetType() == typeof(ComboBox))
                    {
                        newRow[i] = (contrl[0] as ComboBox).SelectedItem;
                    }
                }
                contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][compon_comb["dataGridView" + (b + 1).ToString()].Count() - 1], true);
                newRow[newRow.Table.Columns.Count - 2] = (contrl[0] as ComboBox).SelectedValue;
                contrl = Controls.Find(compon_comb["dataGridView" + (b + 1).ToString()][compon_comb["dataGridView" + (b + 1).ToString()].Count() - 2], true);
                newRow[newRow.Table.Columns.Count - 1] = contrl[0].Text;
                SqlCommandBuilder command = new SqlCommandBuilder(adapter[b]);
                adapter[b].Update(ds[b]);

                adapter[b] = new SqlDataAdapter(String.Format(SQL_query.sql_tables_components, Tables[b], Item[b]), connection);
                ds[b] = new DataSet();
                adapter[b].Fill(ds[b]);
                contrl = Controls.Find("dataGridView" + (b + 1).ToString(), true);
                (contrl[0] as DataGridView).DataSource = ds[b].Tables[0];

            }
        }
    }
}