
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace CompOut
{
    public partial class Form1 : Form
    {
        public static string desktopPath = Directory.GetCurrentDirectory();//���� � exe ���������
        SqlDataAdapter adapter_comps;//������� ���� ������ (������������� ����� ����� �� � ������)
        DataSet ds_comps;//����� ������ ��������
        DataSet ds;//����� ������ ��������
        SqlDataAdapter adapter;//�������� ��
        List<DataSet> Tables;//������ �� ������ ��������
        bool polzovatel = false;//���������� ����������� �� �� ������������� ��� ��������
        int ID_select = 0;//��������� ��������� - id
        string Name_select;// ��� ���������� ����������
        int ID_div ;//��������� �������������
        string Name_div;//��� ����� �������������
        List<Data>[] list = new List<Data>[40];// ������ �������� ����������� � �� id
        Dictionary<string, string[]> compon = new Dictionary<string, string[]>() //������� ������������� ������� � �� �����������
        {
            {"cpu", new string[] {"comboBox3", "comboBox4" } },
            {"mb", new string[] {"comboBox5" } },
            {"memory", new string[] {"comboBox6", "comboBox7", "comboBox8", "comboBox9" } },
            {"video", new string[] {"comboBox10" } },
            {"sound", new string[] {"comboBox11" } },
            {"cases", new string[] {"comboBox12" } },
            {"hdd", new string[] {"comboBox13", "comboBox14", "comboBox15"} },
            {"lan", new string[] {"comboBox16" } },
            {"fdd", new string[] {"comboBox17" } },
            {"cdr", new string[] {"comboBox18" } },
            {"cdrw", new string[] {"comboBox19" } },
            {"dvd", new string[] {"comboBox20" } },
            {"display", new string[] {"comboBox23", "comboBox24" } },
            {"printer", new string[] {"comboBox25", "comboBox26", "comboBox27", "comboBox28" } },
            {"scaner", new string[] {"comboBox29" } },
            {"modem", new string[] {"comboBox30" } },
            {"key", new string[] {"comboBox31" } },
            {"mouse", new string[] {"comboBox32" } },
            {"ups", new string[] {"comboBox33" } }
           // {"usb", new string[] {"comboBox34" } },
           // {"other", new string[] {"comboBox35" } }
        };

        Dictionary<string, string[]> compon_comb = new Dictionary<string, string[]>() //������� ������������� ���������� ���������� � ��� �����
        {
            {"comboBox3", new string[] {"textBox6", "textBox9" } },
            {"comboBox4", new string[] {"textBox7", "textBox8" } },
            {"comboBox5", new string[] {"textBox11"} },
            {"comboBox6", new string[] {"textBox14", "textBox16", "textBox12" } },
            {"comboBox7", new string[] {"textBox13", "textBox15", "textBox10" } },
            {"comboBox8", new string[] {"textBox22", "textBox18", "textBox20" } },
            {"comboBox9", new string[] {"textBox21", "textBox17", "textBox19" } },
            {"comboBox10", new string[] {"textBox24", "textBox23" } },
            {"comboBox11", new string[] {"textBox25"} },
            {"comboBox12", new string[] {"textBox27", "textBox26" } },
            {"comboBox13", new string[] {"textBox31", "textBox29" } },
            {"comboBox14", new string[] {"textBox30", "textBox28" } },
            {"comboBox15", new string[] {"textBox33", "textBox32" } },
            {"comboBox16", new string[] {"textBox35", "textBox34" } },
            {"comboBox17", new string[] {"textBox37", "textBox36" } },

            {"comboBox18", new string[] {"textBox39", "textBox38" } },
            {"comboBox19", new string[] {"textBox41", "textBox40", "textBox42" } },
            {"comboBox20", new string[] {"textBox45", "textBox44", "textBox43", "textBox47", "textBox46" } },

            {"comboBox23", new string[] {"textBox65", "textBox61", "textBox63" } },
            {"comboBox24", new string[] {"textBox64", "textBox60", "textBox62" } },
            {"comboBox25", new string[] { "textBox84", "textBox71",  "textBox69" } },
            {"comboBox26", new string[] { "textBox82", "textBox70",  "textBox68" } },
            {"comboBox27", new string[] { "textBox76", "textBox59",  "textBox57" } },
            {"comboBox28", new string[] { "textBox74", "textBox58", "textBox56" } },
            {"comboBox29", new string[] {"textBox83", "textBox79", "textBox81" } },
            {"comboBox30", new string[] {"textBox73", "textBox72" } },
            {"comboBox31", new string[] {"textBox75"} },
            {"comboBox32", new string[] {"textBox77"} },
            {"comboBox33", new string[] {"textBox80", "textBox78" } }
        };


        public Form1()
        {
            InitializeComponent();                                                          //������ �����
            menu_load();                                                                    //�������� ����
            saveFileDialog1.Filter = "Excel files(*.xls)|*.xls";                            //������ �������� ������

        }

        public void menu_load()
        {
            menu.Nodes.Clear();                                                                 //������ ���� �� ������ �������
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;              //�������� ����������� ������������� ������ �������������
            dataGridView1.AllowUserToAddRows = false;

            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))//������������ � ��
            {
                connection.Open();                                              //��������� �����������
                adapter = new SqlDataAdapter(SQL_query.sql, connection);       //��������� SQL ������

                ds = new DataSet();                                            //������� ������� 
                adapter.Fill(ds);                                               //���������� ������ � ��
                dataGridView1.DataSource = ds.Tables[0];                        //������� ������ � �����
                dataGridView1.RowHeadersVisible = false;                        //�������� ������� � �������� �����

                // 
                dataGridView1.Columns[0].Visible = false;               //������� ������� � id
                dataGridView1.Columns[1].HeaderText = "�������� �������������";  //�������� �������
                dataGridView1.Columns[2].HeaderText = "�.�.�. ����������";
                dataGridView1.Columns[3].HeaderText = "���������� �������";
                dataGridView1.Columns[4].HeaderText = "����������";
                dataGridView1.Columns[1].Width = splitContainer1.Panel2.Width / 3; //������ ��������
                dataGridView1.Columns[2].Width = splitContainer1.Panel2.Width / 4;
                dataGridView1.Columns[3].Width = splitContainer1.Panel2.Width / 6;
                dataGridView1.Columns[4].Width = splitContainer1.Panel2.Width / 4;



                for (int i = 0; i < dataGridView1.Rows.Count; i++) //������� ������������� � ����
                {
                    menu.Nodes.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    menu.Nodes[i].ImageIndex = 11;
                    adapter_comps = new SqlDataAdapter(string.Format(SQL_query.sql_comps, dataGridView1.Rows[i].Cells[0].Value), connection);
                    ds_comps = new DataSet(); //SQL ������ - ������� ����������� ������������ �������������
                    adapter_comps.Fill(ds_comps);
                    if (ds_comps.Tables[0].Rows.Count > 0) //���� � ������������� ���� ����������
                        for (int j = 0; j < ds_comps.Tables[0].Rows.Count; j++)
                        {
                            var cell = ds_comps.Tables[0].Rows[j]; //��������� �� � ��������
                            menu.Nodes[i].Nodes.Add(cell[2].ToString());
                        }
                }
                Load_tables();//��������� ������ �����������
            }
        }


        public void Load_tables()
        {

            Tables = new List<DataSet>();
            for (int i = 0; i < 38; i++)
            {
                list[i] = new List<Data>();//������������� ������ ��� ����������-ID
            }
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open(); //������������ � ��

                for (int i = 0; i < SQL_query.Tables.Count; i++) //������������ � ������ ������� � �������� ������ �� ���
                {
                    var query = string.Format(SQL_query.sql_tables, SQL_query.Tables[i]);
                    adapter = new SqlDataAdapter(query, connection);
                    Tables.Add(new DataSet());
                    adapter.Fill(Tables[i]);
                }
            }
            // ��������� ����������
            //������ - �����
            Control[] ast;//����� ���������� �� �����
            byte b = 0;//�������� �����������
            byte k = 2;
            foreach (DataRow row in Tables[0].Tables[0].Rows) //������� � ������ ������ � �������������
            {
                list[0].Add(new Data(row.Field<int>("div_id"), row.Field<string>("div_name")));
            }

            comboBox1.DataSource = list[0];// ������������� ��������� ������ ����������� �� ������
            comboBox1.DisplayMember = "Name"; // ���������� ��� ����������
            comboBox1.ValueMember = "id"; // � �������� id ����������� ��������
            //������ - �������������
            foreach (DataRow row in Tables[0].Tables[0].Rows)
            {
                list[1].Add(new Data(row.Field<int>("div_id"), row.Field<string>("div_chief")));
            }

            comboBox2.DataSource = list[1];
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "id";
            //���� ����� �� � �����
            foreach (KeyValuePair<string, string[]> s in compon)
            {
                for (int i = 0; i < s.Value.Count(); i++)
                {
                    foreach (DataRow row in Tables[b + 2].Tables[0].Rows)
                    {
                        list[k].Add(new Data(row.Field<int>(SQL_query.Item2[b] + "_id"), row.Field<string>(SQL_query.Item2[b] + "_model")));
                    }
                    ast = Controls.Find(s.Value[i], true);
                    (ast[0] as ComboBox).DataSource = list[k];//��������� ������� � ���������� ���� �����������
                    (ast[0] as ComboBox).DisplayMember = "Name";
                    (ast[0] as ComboBox).ValueMember = "id";
                    k++;
                }
                b++;
            }
            //��������� ����
            foreach (DataRow row in Tables[21].Tables[0].Rows)
            {
                list[31].Add(new Data(row.Field<int>("soft_id"), row.Field<string>("soft_name")));
            }

            comboBox22.DataSource = list[31];
            comboBox22.DisplayMember = "Name";
            comboBox22.ValueMember = "id";
            //������ - �������
            foreach (DataRow row in Tables[22].Tables[0].Rows)
            {
                list[32].Add(new Data(row.Field<int>("master_id"), row.Field<string>("master_name")));
            }

            comboBox21.DataSource = list[32];
            comboBox21.DisplayMember = "Name";
            comboBox21.ValueMember = "id";
            //������ ����������
            foreach (DataRow row in Tables[23].Tables[0].Rows)
            {
                list[33].Add(new Data(row.Field<int>("other_id"), row.Field<string>("other_model")));
            }

            comboBox35.DataSource = list[33];
            comboBox35.DisplayMember = "Name";
            comboBox35.ValueMember = "id";
            // usb ����������
            foreach (DataRow row in Tables[25].Tables[0].Rows)
            {
                list[34].Add(new Data(row.Field<int>("usb_id"), row.Field<string>("usb_model")));
            }

            comboBox34.DataSource = list[34];
            comboBox34.DisplayMember = "Name";
            comboBox34.ValueMember = "id";


        }



        private void dataGridView1_SelectionChanged(object sender, EventArgs e) //��� ������ ������������� 
        {
            if (dataGridView1.SelectedCells.Count != 0) //���� ������� �������������   ��
            {
                DataGridViewSelectedCellCollection DGVCell = dataGridView1.SelectedCells; //�������� ����� ���������� �������
                var dgvc = DGVCell[1];//���������� ������ �� ���������� ������
                var dgvr = dgvc.OwningRow;

                name_podraz_textbox.Text = dgvr.Cells[1].Value.ToString();//��������� ���� ��� ��������������
                fio_chef_textbox.Text = dgvr.Cells[2].Value.ToString();
                phone_textbox.Text = dgvr.Cells[3].Value.ToString();
                prim_textbox.Text = dgvr.Cells[4].Value.ToString();
                ID_div = int.Parse(dgvr.Cells[0].Value.ToString());
                Name_div = name_podraz_textbox.Text;
            }
        }

        private void Delete_button_Click(object sender, EventArgs e) //������ ��������
        {
            if (polzovatel)//���� ��������� � �������������� ������������
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.sql_comps_delete, ID_select); //��������� SQL ������ �� ��������
                    adapter = new SqlDataAdapter(query, connection);

                    ds = new DataSet();
                    adapter.Fill(ds);
                    menu_load(); //��������� ����
                    menu.SelectedNode = menu.Nodes[0]; 
                }
            }
            else     //���� ��������� � ���� �������������
            {
                int index = dataGridView1.CurrentCell.RowIndex; // ����� ���������� �������
                string id = dataGridView1.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.Delete_button, id);//SQL ������ �� �������� �������������
                    adapter = new SqlDataAdapter(query, connection);

                    ds = new DataSet();
                    adapter.Fill(ds);

                    adapter = new SqlDataAdapter(SQL_query.sql, connection);

                    ds = new DataSet();
                    adapter.Fill(ds);//��������� ������ � ���� � �������
                    dataGridView1.DataSource = ds.Tables[0];
                    menu_load();
                }
            }
        }

        private void Paste_button_Click(object sender, EventArgs e)//������ �������
        {

            if (Paste_button.Text.Equals("��������")) // ���� ������ � 1 ���
            {
                if (polzovatel) //���� ����������� �����������
                {
                    clear_info(); //������� ���
                    Edit_button.Enabled = false; //������ ������������ ������ ������
                    Delete_button.Enabled = false;
                    Paste_button.Text = "���������";
                }
                else//���� ��������� �������������
                {
                    Edit_button.Enabled = false;//������ ������������ ������ ������
                    Delete_button.Enabled = false;
                    Paste_button.Text = "���������";
                    name_podraz_textbox.Text = "";//������� ���
                    fio_chef_textbox.Text = "";
                    phone_textbox.Text = "";
                    prim_textbox.Text = "";
                }
            }
            else //���� ������ ������ ���
            {
                if (polzovatel) //� ������ ������������
                {
                    if (!textBox1.Text.Equals("") && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)//���� ��������� �������� ����
                    {
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            adapter_comps = new SqlDataAdapter(SQL_query.sql_comps_all, connection);
                            ds_comps = new DataSet();  //������������ � ������� �������������
                            adapter_comps.Fill(ds_comps);
                            DataTable dt = ds_comps.Tables[0]; //�� �������� �������� �������
                            DataRow newRow = dt.NewRow();//������� ����� ������� � �������
                            newRow[1] = int.Parse(comboBox1.SelectedValue.ToString());//�������� ��������� �������
                            newRow[2] = textBox1.Text;
                            newRow[3] = textBox4.Text;
                            newRow[4] = textBox2.Text;
                            newRow[5] = textBox3.Text;
                            newRow[6] = dateTimePicker1.Value.ToShortDateString();
                            newRow[7] = textBox5.Text;

                            Control[] ast;
                            int k = 8;
                            foreach (KeyValuePair<string, string[]> s in compon) //���������� ���������� ����������� �������� � ��� ������
                            {
                                for (int i = 0; i < s.Value.Count(); i++)
                                {
                                    ast = Controls.Find(s.Value[i], true);
                                    if ((ast[0] as ComboBox).SelectedIndex != null && (ast[0] as ComboBox).SelectedIndex != -1)
                                    {
                                        newRow[k] = int.Parse((ast[0] as ComboBox).SelectedValue.ToString());
                                    }
                                    k++;
                                }
                            }
                            newRow[k] = textBox67.Text; k++;//������� �� ���������
                            newRow[k] = textBox66.Text; k++;
                            newRow[k] = textBox54.Text; k++;
                            newRow[k] = textBox53.Text; k++;
                            dt.Rows.Add(newRow);//��������� ������������ �������
                            SqlCommandBuilder command = new SqlCommandBuilder(adapter_comps);//��������� ������� �������
                            adapter_comps.Update(ds_comps);//��������� �������

                            menu_load();//��������� ����
                            Edit_button.Enabled = true;//��������� ������ � ��������� ���������
                            Delete_button.Enabled = true;
                            Paste_button.Text = "��������";
                            menu.SelectedNode = menu.Nodes[0];
                        }
                    }
                    else //���� ��������� �� ���������
                    {
                        MessageBox.Show("������ ����� ������"); //������ ������
                        Edit_button.Enabled = true;
                        Delete_button.Enabled = true;
                        Paste_button.Text = "��������";
                        menu.SelectedNode = menu.Nodes[0];
                    }
                }
                else
                {// ���� ��������� �������������
                    if (fio_chef_textbox.Text.Equals("") && phone_textbox.Text.Equals("") && name_podraz_textbox.Text.Equals(""))
                    {
                        MessageBox.Show("������ ����� ������"); //������ ���������� �������� �����
                        Edit_button.Enabled = true;
                        Delete_button.Enabled = true;
                        Paste_button.Text = "��������";
                        menu.Select();
                    }
                    else
                    {// ���� ��������� ���������
                        Edit_button.Enabled = true;
                        Delete_button.Enabled = true;
                        Paste_button.Text = "��������";
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            connection.Open(); // ��������� SQL ������ �� �������
                            var query = string.Format(SQL_query.Save_button, name_podraz_textbox.Text, fio_chef_textbox.Text, phone_textbox.Text, prim_textbox.Text);
                            adapter = new SqlDataAdapter(query, connection);

                            ds = new DataSet();
                            adapter.Fill(ds);

                            adapter = new SqlDataAdapter(SQL_query.sql, connection);

                            ds = new DataSet();//��������� ������� � ���������������
                            adapter.Fill(ds);//��������� ����� ������� �� �����
                            dataGridView1.DataSource = ds.Tables[0];
                        }
                    }
                }
            }
        }

        private void Edit_button_Click(object sender, EventArgs e)//������ ��������������
        {
            if (polzovatel)
            {
                if (!textBox1.Text.Equals("") && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
                {
                    using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                    {
                        adapter_comps = new SqlDataAdapter(SQL_query.sql_comps_all, connection);
                        ds_comps = new DataSet();
                        adapter_comps.Fill(ds_comps);
                        DataTable dt = ds_comps.Tables[0];
                        DataRow newRow = dt.Select("comp_id = " + ID_select.ToString())[0]; //���������� �������, ������ �� ������� � �������� ���������� ������
                        newRow[1] = int.Parse(comboBox1.SelectedValue.ToString());
                        newRow[2] = textBox1.Text;
                        newRow[3] = textBox4.Text;
                        newRow[4] = textBox2.Text;
                        newRow[5] = textBox3.Text;
                        newRow[6] = dateTimePicker1.Value.ToShortDateString();
                        newRow[7] = textBox5.Text;

                        Control[] ast;
                        int k = 8;
                        foreach (KeyValuePair<string, string[]> s in compon) //���������� ��� ����������
                        {
                            for (int i = 0; i < s.Value.Count(); i++)
                            {
                                ast = Controls.Find(s.Value[i], true);
                                if ((ast[0] as ComboBox).SelectedIndex != null && (ast[0] as ComboBox).SelectedIndex != -1)
                                {
                                    newRow[k] = int.Parse((ast[0] as ComboBox).SelectedValue.ToString());
                                }
                                k++;
                            }
                        }
                        newRow[k] = textBox67.Text; k++; //���������� ������� ���� ���������
                        newRow[k] = textBox66.Text; k++;
                        newRow[k] = textBox54.Text; k++;
                        newRow[k] = textBox53.Text; k++;
                        SqlCommandBuilder command = new SqlCommandBuilder(adapter_comps);
                        adapter_comps.Update(ds_comps);

                        menu_load(); //�������� ���� � ������
                        Edit_button.Enabled = true;
                        Delete_button.Enabled = true;
                        Paste_button.Text = "��������";
                        menu.SelectedNode = menu.Nodes[0];
                    }
                }
            }
            else
            {//���� ����������� �������������
                int index = dataGridView1.CurrentCell.RowIndex; // ����� ���������� �������
                string id = dataGridView1.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open(); // ��������� SQL ������ �� ��������������
                    var query = string.Format(SQL_query.Edit_button, name_podraz_textbox.Text, fio_chef_textbox.Text, phone_textbox.Text, prim_textbox.Text, id);
                    adapter = new SqlDataAdapter(query, connection);

                    ds = new DataSet();
                    adapter.Fill(ds);

                    adapter = new SqlDataAdapter(SQL_query.sql, connection);

                    ds = new DataSet();//��������� ������� 
                    adapter.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
                menu_load();
            }
        }

        private void menu_AfterSelect(object sender, TreeViewEventArgs e)
        {//����� ������ ����
            if (e.Node.Level == 0)   //������� ������� ����
            {
                dataGridView1.ClearSelection();// �������� ������� � �������, � ����������� �� ���������� �������������
                dataGridView1.Rows[e.Node.Index].Selected = true;
                polzovatel = false;
                panel4.Visible = false;
                panel3.Visible = true;
                label1.Visible = true;
                dataGridView1.Visible = true;
            }
            else
            if (e.Node.Level == 1)   //��������� ������� ����
            {
                panel3.Visible = false; // ������� ���������� � ��������
                label1.Visible = false;
                dataGridView1.Visible = false;
                panel4.Visible = true;
                panel4.Dock = DockStyle.Fill;
                polzovatel = true;
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    adapter = new SqlDataAdapter(string.Format(SQL_query.Comps_info, e.Node.Text), connection);
                    ds = new DataSet();//��������� ���������� �� ���������� � ���� �������� �� ��
                    adapter.Fill(ds);
                    var strok = ds.Tables[0].Rows[0];
                    ID_select = (int)strok[0];
                }
                Name_select = e.Node.Text;
                obnovlenie_inf(e.Node.Text); //��������� ��������� �����������

            }
        }

        private void obnovlenie_inf(string name) //���������� ����������� ��� ������ ����������
        {
            clear_info();//������� ������ ������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();//SQL ������ �������� � ���������� � �����
                var query = string.Format(SQL_query.Comps_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                var cell = ds.Tables[0].Rows[0];
                textBox1.Text = cell[2].ToString();
                textBox2.Text = cell[4].ToString();
                textBox3.Text = cell[5].ToString();
                textBox4.Text = cell[3].ToString();
                textBox5.Text = cell[7].ToString();
                query = string.Format(SQL_query.sql_where, cell[1]);//SQL ������ �������� � �������������
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                var div = ds.Tables[0].Rows[0];
                var index = comboBox1.FindString(div[1].ToString());
                comboBox1.SelectedIndex = index;
                index = comboBox2.FindString(div[2].ToString());
                comboBox2.SelectedIndex = index;

                Control[] ast;

                for (int i = 8; i < 26; i++)//���������� ��������� ������� � ������������
                {

                    ast = Controls.Find("comboBox" + (i - 5).ToString(), true);
                    if (DBNull.Value.Equals(cell[i])) (ast[0] as ComboBox).SelectedIndex = -1;  //��������� ��� �� ������ ��������
                    else
                        if (int.Parse(cell[i].ToString()) == 0) (ast[0] as ComboBox).SelectedIndex = -1; // ��������� ��� �� ����� 0 (������� ���� � 1)
                    else
                    {
                        query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[i - 6], SQL_query.Item[i - 6], cell[i].ToString());
                        adapter = new SqlDataAdapter(query, connection);//����� �� �� ������������� �����������
                        ds = new DataSet();
                        adapter.Fill(ds);
                        //���������
                        var cel = ds.Tables[0].Rows[0];
                        index = (ast[0] as ComboBox).FindString(cel[1].ToString());
                        (ast[0] as ComboBox).SelectedIndex = index;
                        var s = compon_comb[ast[0].Name];
                        ast = Controls.Find(s[s.Count() - 1], true);
                        ast[0].Text = cel[cel.ItemArray.Count() - 1].ToString();//������� �������� � ����������� � ���������� � ��������� ����
                        if (s.Count() > 1)
                        {
                            for (int j = 0; j < s.Count() - 1; j++)
                            {
                                ast = Controls.Find(s[j], true);
                                ast[0].Text = cel[2 + j].ToString();
                            }
                        }
                    }
                }
                for (int i = 26; i < 37; i++)
                {

                    ast = Controls.Find("comboBox" + (i - 3).ToString(), true);
                    if (DBNull.Value.Equals(cell[i])) (ast[0] as ComboBox).SelectedIndex = -1;  //��������� ��� �� ������ ��������
                    else
                        if (int.Parse(cell[i].ToString()) == 0) (ast[0] as ComboBox).SelectedIndex = -1; // ��������� ��� �� ����� 0 (������� ���� � 1)
                    else
                    {
                        query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[i - 6], SQL_query.Item[i - 6], cell[i].ToString());
                        adapter = new SqlDataAdapter(query, connection);
                        ds = new DataSet();
                        adapter.Fill(ds);
                        //���������
                        var cel = ds.Tables[0].Rows[0];
                        index = (ast[0] as ComboBox).FindString(cel[1].ToString());
                        (ast[0] as ComboBox).SelectedIndex = index;
                        var s = compon_comb[ast[0].Name];
                        ast = Controls.Find(s[s.Count() - 1], true);
                        ast[0].Text = cel[cel.ItemArray.Count() - 1].ToString();
                        if (s.Count() > 1)
                        {
                            for (int j = 0; j < s.Count() - 1; j++)
                            {
                                ast = Controls.Find(s[j], true);
                                ast[0].Text = cel[2 + j].ToString();
                            }
                        }
                    }
                }
                //�������� � ���������
                if (DBNull.Value.Equals(cell[37])) textBox67.Text = "";  //��������� ��� �� ������ ��������
                else
                {
                    textBox67.Text = cell[37].ToString();
                }

                if (DBNull.Value.Equals(cell[38])) textBox66.Text = "";  //��������� ��� �� ������ ��������
                else
                {
                    textBox66.Text = cell[38].ToString();
                }

                if (DBNull.Value.Equals(cell[39])) textBox54.Text = "";  //��������� ��� �� ������ ��������
                else
                {
                    textBox54.Text = cell[39].ToString();
                }

                if (DBNull.Value.Equals(cell[40])) textBox53.Text = "";  //��������� ��� �� ������ ��������
                else
                {
                    textBox53.Text = cell[40].ToString();
                }

                //��������� � ������� �������� � ��, ���, �����
                query = string.Format(SQL_query.usb_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView5.DataSource = ds.Tables[0];
                dataGridView5.Columns[0].Visible = false;
                dataGridView5.Columns[1].Width = splitContainer1.Panel2.Width / 3;
                dataGridView5.Columns[2].Width = splitContainer1.Panel2.Width / 3;
                dataGridView5.Columns[3].Width = splitContainer1.Panel2.Width / 4;
                dataGridView5.Columns[1].HeaderText = "��������";
                dataGridView5.Columns[2].HeaderText = "�������������";
                dataGridView5.Columns[3].HeaderText = "����������";
                dataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView5.AllowUserToAddRows = false;


                query = string.Format(SQL_query.other_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
                dataGridView6.Columns[0].Visible = false;
                dataGridView6.Columns[1].Width = splitContainer1.Panel2.Width / 3;
                dataGridView6.Columns[2].Width = splitContainer1.Panel2.Width / 3;
                dataGridView6.Columns[3].Width = splitContainer1.Panel2.Width / 4;
                dataGridView6.Columns[1].HeaderText = "��������";
                dataGridView6.Columns[2].HeaderText = "�������������";
                dataGridView6.Columns[3].HeaderText = "����������";
                dataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView6.AllowUserToAddRows = false;


                query = string.Format(SQL_query.soft_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView4.DataSource = ds.Tables[0];
                dataGridView4.Columns[0].Visible = false;
                dataGridView4.Columns[1].Width = splitContainer1.Panel2.Width / 2;
                dataGridView4.Columns[2].Width = splitContainer1.Panel2.Width / 4;
                dataGridView4.Columns[3].Width = splitContainer1.Panel2.Width / 4;
                dataGridView4.Columns[1].HeaderText = "��������";
                dataGridView4.Columns[2].HeaderText = "�������������";
                dataGridView4.Columns[3].HeaderText = "����������";
                dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView4.AllowUserToAddRows = false;

                query = string.Format(SQL_query.service_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView3.DataSource = ds.Tables[0];
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Width = splitContainer1.Panel2.Width / 2;
                dataGridView3.Columns[2].Width = splitContainer1.Panel2.Width / 4;
                dataGridView3.Columns[3].Width = splitContainer1.Panel2.Width / 4;
                dataGridView3.Columns[1].HeaderText = "�.�.� �������";
                dataGridView3.Columns[2].HeaderText = "����";
                dataGridView3.Columns[3].HeaderText = "����������";
                dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView3.AllowUserToAddRows = false;


                query = string.Format(SQL_query.user_info, name);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[1].Width = splitContainer1.Panel2.Width / 4;
                dataGridView2.Columns[2].Width = splitContainer1.Panel2.Width / 4;
                dataGridView2.Columns[3].Width = splitContainer1.Panel2.Width / 4;
                dataGridView2.Columns[4].Width = splitContainer1.Panel2.Width / 4;
                dataGridView2.Columns[1].HeaderText = "�.�.� ������������";
                dataGridView2.Columns[2].HeaderText = "�����";
                dataGridView2.Columns[3].HeaderText = "������";
                dataGridView2.Columns[4].HeaderText = "����������";
                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView2.AllowUserToAddRows = false;

            }
        }

        public void clear_info()
        {
            dataGridView2.DataSource = "";//������ ��� �������
            dataGridView3.DataSource = "";
            dataGridView4.DataSource = "";
            dataGridView5.DataSource = "";
            dataGridView6.DataSource = "";
            Control[] ast;
            textBox1.Text = ""; //��� ���� ������
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            foreach (KeyValuePair<string, string[]> s in compon) //��� ����������
            {
                for (int i = 0; i < s.Value.Count(); i++)
                {
                    ast = Controls.Find(s.Value[i], true);
                    (ast[0] as ComboBox).SelectedIndex = -1;
                }
            }

            foreach (KeyValuePair<string, string[]> s in compon_comb)
            {
                for (int i = 0; i < s.Value.Count(); i++)
                {
                    ast = Controls.Find(s.Value[i], true);
                    ast[0].Text = "";
                }
            }
        }
        //�����
        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        //�������� ����
        private void button9_Click(object sender, EventArgs e)
        {
            if (comboBox22.SelectedIndex != -1)
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.New_comp_soft, ID_select, comboBox22.SelectedValue);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.soft_info, Name_select);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];
                }

            }
        }
        //������� ����
        private void button11_Click(object sender, EventArgs e)
        {
            int index = dataGridView4.CurrentCell.RowIndex; // ����� ���������� �������
            string id = dataGridView4.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                var query = string.Format(SQL_query.Delete_soft, id);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                query = string.Format(SQL_query.soft_info, Name_select);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView4.DataSource = ds.Tables[0];
            }
        }
        //�������� ������
        private void button12_Click(object sender, EventArgs e)
        {
            if (comboBox21.SelectedIndex != -1)
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    string query = string.Format(SQL_query.sql_tables, "service");
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    DataTable dt = ds.Tables[0];
                    DataRow newRow = dt.NewRow();
                    newRow[1] = ID_select;
                    newRow[2] = comboBox21.SelectedValue;
                    newRow[3] = dateTimePicker2.Value.ToShortDateString();
                    newRow[4] = textBox55.Text;
                    dt.Rows.Add(newRow);
                    SqlCommandBuilder command = new SqlCommandBuilder(adapter);
                    adapter.Update(ds);
                    query = string.Format(SQL_query.service_info, Name_select);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView3.DataSource = ds.Tables[0];
                }

            }
        }
        //������� ������
        private void button10_Click(object sender, EventArgs e)
        {
            int index = dataGridView3.CurrentCell.RowIndex; // ����� ���������� �������
            string id = dataGridView3.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                var query = string.Format(SQL_query.Delete_service, id);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                query = string.Format(SQL_query.service_info, Name_select);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView3.DataSource = ds.Tables[0];
            }
        }
        //�������� ������������
        private void button14_Click(object sender, EventArgs e)
        {
            if (!textBox50.Text.Equals("") && !textBox49.Text.Equals("") && !textBox48.Text.Equals(""))
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.New_user, ID_select, textBox50.Text, textBox49.Text, textBox48.Text, textBox51.Text);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.user_info, Name_select);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];
                }

            }
        }
        //������� ������������
        private void button13_Click(object sender, EventArgs e)
        {
            int index = dataGridView2.CurrentCell.RowIndex; // ����� ���������� �������
            string id = dataGridView2.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                var query = string.Format(SQL_query.Delete_user, id);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                query = string.Format(SQL_query.user_info, Name_select);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
            }
        }
        //�������� usb
        private void button16_Click(object sender, EventArgs e)
        {
            if (comboBox34.SelectedIndex != -1)
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.New_usb, ID_select, comboBox34.SelectedValue);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.usb_info, Name_select);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];
                }

            }
        }
        //������� usb
        private void button15_Click(object sender, EventArgs e)
        {
            int index = dataGridView5.CurrentCell.RowIndex; // ����� ���������� �������
            string id = dataGridView5.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                var query = string.Format(SQL_query.Delete_usb, id);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                query = string.Format(SQL_query.usb_info, Name_select);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView5.DataSource = ds.Tables[0];
            }
        }
        //�������� ������ ����������
        private void button18_Click(object sender, EventArgs e)
        {
            if (comboBox35.SelectedIndex != -1)
            {
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    connection.Open();
                    var query = string.Format(SQL_query.New_other, ID_select, comboBox35.SelectedValue);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    query = string.Format(SQL_query.other_info, Name_select);
                    adapter = new SqlDataAdapter(query, connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];
                }

            }
        }
        //������� ������ ����������
        private void button17_Click(object sender, EventArgs e)
        {
            int index = dataGridView6.CurrentCell.RowIndex; // ����� ���������� �������
            string id = dataGridView6.Rows[index].Cells[0].Value.ToString(); // �������� id ���������� �������
            using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
            {
                connection.Open();
                var query = string.Format(SQL_query.Delete_other, id);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                query = string.Format(SQL_query.other_info, Name_select);
                adapter = new SqlDataAdapter(query, connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
            }
        }
        //��������� �������� ������-���� ����������
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox c = (ComboBox)sender;
            if (c.SelectedIndex >= 0 && c.SelectedValue.GetType() == typeof(int))
            {
                Control[] contrl;
                int b = int.Parse((string)((ComboBox)sender).Tag);
                var r = c.SelectedValue;
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    adapter = new SqlDataAdapter(String.Format(SQL_query.sql_tables_where, SQL_query.Tables2[b], SQL_query.Item[b], r), connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    DataTable dt = ds.Tables[0];
                    DataRow newRow = dt.Rows[0];
                    for (int i = 0; i < compon_comb["comboBox" + (b + 1).ToString()].Count(); i++)
                    {
                        contrl = Controls.Find(compon_comb["comboBox" + (b + 1).ToString()][i], true);
                        contrl[0].Text = newRow[i + 2].ToString();

                    }

                    contrl = Controls.Find(compon_comb["comboBox" + (b + 1).ToString()][compon_comb["comboBox" + (b + 1).ToString()].Count() - 1], true);
                    contrl[0].Text = newRow[newRow.Table.Columns.Count - 1].ToString();

                }
            }

        }
        Form_master fr_m;//����� � ���������
        Form_proizv fr_p;//����� � ���������������
        Form_compon fr_c;//����� � ������������
        public static Form IsFormAlreadyOpen(Type FormType)
        {
            foreach (Form OpenForm in System.Windows.Forms.Application.OpenForms) //���� ��������� ��������� ����� ������, �� ������ �� ������
            { //���� �� ������ - �������
                if (OpenForm.GetType() == FormType)
                    return OpenForm;
            }

            return null;
        }
        private void button4_Click(object sender, EventArgs e) //�������� ����� � ���������
        {
            if ((fr_p = (Form_proizv)IsFormAlreadyOpen(typeof(Form_proizv))) == null)
            { //Form isn't open so create one
                fr_p = new Form_proizv();
                fr_p.Show();
            }
            else
            { // Form is already open so bring it to the front
                fr_p.BringToFront();

            }
        }

        private void button3_Click(object sender, EventArgs e)//�������� ����� � ���������������
        {
            if ((fr_c = (Form_compon)IsFormAlreadyOpen(typeof(Form_compon))) == null)
            { //Form isn't open so create one
                fr_c = new Form_compon();
                fr_c.Show();
            }
            else
            { // Form is already open so bring it to the front
                fr_m.BringToFront();

            }

        }

        private void button2_Click(object sender, EventArgs e) //�������� ����� � ������������
        {
            if ((fr_m = (Form_master)IsFormAlreadyOpen(typeof(Form_master))) == null)
            { //Form isn't open so create one
                fr_m = new Form_master();
                fr_m.Show();
            }
            else
            { // Form is already open so bring it to the front
                fr_m.BringToFront();

            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //��������� ���������� ���������
        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox c = (ComboBox)sender;
            if (c.SelectedIndex >= 0 && c.SelectedValue.GetType() == typeof(int))
            {
                Control[] contrl;
                int b = int.Parse((string)((ComboBox)sender).Tag);
                var r = c.SelectedValue;
                using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                {
                    adapter = new SqlDataAdapter(String.Format(SQL_query.sql_tables_where, SQL_query.Tables2[b], SQL_query.Item[b], r), connection);
                    ds = new DataSet();
                    adapter.Fill(ds);
                    DataTable dt = ds.Tables[0];
                    DataRow newRow = dt.Rows[0];
                    for (int i = 0; i < compon_comb["comboBox" + (b + 3).ToString()].Count(); i++)
                    {
                        contrl = Controls.Find(compon_comb["comboBox" + (b + 3).ToString()][i], true);
                        contrl[0].Text = newRow[i + 2].ToString();

                    }

                    contrl = Controls.Find(compon_comb["comboBox" + (b + 3).ToString()][compon_comb["comboBox" + (b + 3).ToString()].Count() - 1], true);
                    contrl[0].Text = newRow[newRow.Table.Columns.Count - 1].ToString();

                }

            }
        }

        //���������� �������
        Excel.Application? app;
        private void button5_Click(object sender, EventArgs e)
        {
            if (ID_select > 0)
            {
                DialogResult dialogResult = MessageBox.Show("��������� �������� ����������: " + Name_select, "�������� ����������", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        app =  new Excel.Application();//��������� exeL
                        Excel.Workbook xlWB = app.Workbooks.Open(desktopPath + @"\docs\Card.xls", //��������� ������
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
                        Excel.Worksheet xlSht = xlWB.Worksheets["����1"];

                        // �������� ���������
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            connection.Open();
                            var query = string.Format(SQL_query.Comps_info, Name_select);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var cell = ds.Tables[0].Rows[0];
                            xlSht.Cells[4, 2] = cell[2].ToString();  //��� ����������
                            DateTime dt = (DateTime) cell[6];
                            xlSht.Cells[5, 2] = dt.ToShortDateString(); //���� ����� � ����
                            xlSht.Cells[3, 4] = cell[5].ToString();  //�������
                            xlSht.Cells[4, 4] = cell[4].ToString();  //ip �����
                            xlSht.Cells[5, 4] = cell[3].ToString();  //���������� �����

                            query = string.Format(SQL_query.sql_where, cell[1]);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var div = ds.Tables[0].Rows[0];
                            xlSht.Cells[3, 2] = div[1].ToString();
                            var cel = ds.Tables[0].Rows[0];
                            if (!DBNull.Value.Equals(cell[8]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[2], SQL_query.Item[2], cell[8].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[10, 2] = cel[1].ToString();  //��������� 1
                                xlSht.Cells[11, 2] = cel[2].ToString();  //�������� �������
                            }
                            if (!DBNull.Value.Equals(cell[9]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[3], SQL_query.Item[3], cell[9].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[10, 4] = cel[1].ToString();  //��������� 2
                                xlSht.Cells[11, 4] = cel[2].ToString();  //�������� �������
                            }
                            if (!DBNull.Value.Equals(cell[10]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[4], SQL_query.Item[4], cell[10].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[7, 2] = cel[1].ToString();  //����������� �����
                            }
                            if (!DBNull.Value.Equals(cell[11]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[5], SQL_query.Item[5], cell[11].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[13, 2] = cel[1].ToString();  //����������� ������
                                xlSht.Cells[14, 2] = cel[2].ToString();
                                xlSht.Cells[15, 2] = cel[3].ToString();

                            }
                            if (!DBNull.Value.Equals(cell[12]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[6], SQL_query.Item[6], cell[12].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[13, 4] = cel[1].ToString();  //����������� ������
                                xlSht.Cells[14, 4] = cel[2].ToString();
                                xlSht.Cells[15, 4] = cel[3].ToString();

                            }
                            if (!DBNull.Value.Equals(cell[13]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[7], SQL_query.Item[7], cell[13].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[13, 6] = cel[1].ToString();  //����������� ������
                                xlSht.Cells[14, 6] = cel[2].ToString();
                                xlSht.Cells[15, 6] = cel[3].ToString();

                            }
                            if (!DBNull.Value.Equals(cell[14]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[8], SQL_query.Item[8], cell[14].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[13, 8] = cel[1].ToString();  //����������� ������
                                xlSht.Cells[14, 8] = cel[2].ToString();
                                xlSht.Cells[15, 8] = cel[3].ToString();

                            }
                            if (!DBNull.Value.Equals(cell[15]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[9], SQL_query.Item[9], cell[15].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[20, 2] = cel[1].ToString();  //����������
                                xlSht.Cells[20, 4] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[16]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[10], SQL_query.Item[10], cell[16].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[22, 2] = cel[1].ToString();  //�������� �����
                            }
                            if (!DBNull.Value.Equals(cell[17]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[11], SQL_query.Item[11], cell[17].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[24, 2] = cel[1].ToString();  //������
                                xlSht.Cells[24, 4] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[18]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[12], SQL_query.Item[12], cell[18].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[17, 2] = cel[1].ToString();  //�������  ����
                                xlSht.Cells[18, 2] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[19]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[13], SQL_query.Item[13], cell[19].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[17, 4] = cel[1].ToString();  //�������  ����
                                xlSht.Cells[18, 4] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[20]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[14], SQL_query.Item[14], cell[20].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[17, 6] = cel[1].ToString();  //�������  ����
                                xlSht.Cells[18, 6] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[21]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[15], SQL_query.Item[15], cell[21].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[35, 2] = cel[1].ToString();  //������� �����
                                xlSht.Cells[35, 4] = cel[2].ToString();
                            }

                            if (!DBNull.Value.Equals(cell[22]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[16], SQL_query.Item[16], cell[22].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[33, 2] = cel[1].ToString();  //fdd
                                xlSht.Cells[33, 4] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[23]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[17], SQL_query.Item[17], cell[23].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[26, 2] = cel[1].ToString();  //��-���
                                xlSht.Cells[26, 4] = cel[2].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[24]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[18], SQL_query.Item[18], cell[24].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[28, 2] = cel[1].ToString();  //��-��
                                xlSht.Cells[28, 4] = cel[2].ToString();
                                xlSht.Cells[28, 6] = cel[3].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[25]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[19], SQL_query.Item[19], cell[25].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[30, 2] = cel[1].ToString();  //���
                                xlSht.Cells[30, 4] = cel[2].ToString();
                                xlSht.Cells[30, 6] = cel[3].ToString();
                                xlSht.Cells[31, 4] = cel[4].ToString();
                                xlSht.Cells[31, 6] = cel[5].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[26]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[20], SQL_query.Item[20], cell[26].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[37, 2] = cel[1].ToString();  //�������
                                xlSht.Cells[38, 2] = cel[2].ToString();
                                xlSht.Cells[39, 2] = cel[3].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[27]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[21], SQL_query.Item[21], cell[27].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[37, 4] = cel[1].ToString();  //�������
                                xlSht.Cells[38, 4] = cel[2].ToString();
                                xlSht.Cells[39, 4] = cel[3].ToString();
                            }

                            if (!DBNull.Value.Equals(cell[28]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[22], SQL_query.Item[22], cell[28].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[41, 2] = cel[1].ToString();  //�������
                                xlSht.Cells[42, 2] = cell[37].ToString();
                                xlSht.Cells[43, 2] = cel[2].ToString();
                                xlSht.Cells[44, 2] = cel[3].ToString();
                                xlSht.Cells[45, 2] = cel[4].ToString();
                            }
                            if (!DBNull.Value.Equals(cell[29]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[23], SQL_query.Item[23], cell[29].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[41, 4] = cel[1].ToString();  //�������
                                xlSht.Cells[42, 4] = cell[38].ToString();
                                xlSht.Cells[43, 4] = cel[2].ToString();
                                xlSht.Cells[44, 4] = cel[3].ToString();
                                xlSht.Cells[45, 4] = cel[4].ToString();

                            }

                            if (!DBNull.Value.Equals(cell[30]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[24], SQL_query.Item[24], cell[30].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[41, 6] = cel[1].ToString();  //�������
                                xlSht.Cells[42, 6] = cell[38].ToString();
                                xlSht.Cells[43, 6] = cel[2].ToString();
                                xlSht.Cells[44, 6] = cel[3].ToString();
                                xlSht.Cells[45, 6] = cel[4].ToString();

                            }
                            if (!DBNull.Value.Equals(cell[31]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[25], SQL_query.Item[25], cell[31].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[41, 8] = cel[1].ToString();  //�������
                                xlSht.Cells[42, 8] = cell[38].ToString();
                                xlSht.Cells[43, 8] = cel[2].ToString();
                                xlSht.Cells[44, 8] = cel[3].ToString();
                                xlSht.Cells[45, 8] = cel[4].ToString();

                            }

                            if (!DBNull.Value.Equals(cell[32]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[26], SQL_query.Item[26], cell[32].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[47, 2] = cel[1].ToString();  //������
                                xlSht.Cells[47, 4] = cel[2].ToString();
                                xlSht.Cells[47, 6] = cel[3].ToString();
                            }

                            if (!DBNull.Value.Equals(cell[33]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[27], SQL_query.Item[27], cell[33].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[51, 2] = cel[1].ToString();  //�����
                                xlSht.Cells[51, 4] = cel[2].ToString();
                            }

                            if (!DBNull.Value.Equals(cell[34]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[28], SQL_query.Item[28], cell[34].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[53, 2] = cel[1].ToString();  //����������
                            }

                            if (!DBNull.Value.Equals(cell[35]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[29], SQL_query.Item[29], cell[35].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[55, 2] = cel[1].ToString();  //����
                            }
                            if (!DBNull.Value.Equals(cell[36]))
                            {
                                query = string.Format(SQL_query.sql_tables_where, SQL_query.Tables2[30], SQL_query.Item[30], cell[36].ToString());
                                adapter = new SqlDataAdapter(query, connection);
                                ds = new DataSet();
                                adapter.Fill(ds);
                                cel = ds.Tables[0].Rows[0];
                                xlSht.Cells[49, 2] = cel[1].ToString();  //���
                                xlSht.Cells[49, 4] = cel[2].ToString();  //���
                            }
                            Excel.Range range2 = xlSht.get_Range("A1", "H55");
                            range2.EntireColumn.AutoFit();
                        }
                        if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                            return;
                        // �������� ��������� ����
                        string filename = saveFileDialog1.FileName;
                        app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlSht = null;
                        xlWB = null;
                        app.Visible = true;
                    }
                    finally
                    {
                        

                    }
                }
            }
            else
            {
                MessageBox.Show("�������� ����������");
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (ID_select > 0)
            {
                DialogResult dialogResult = MessageBox.Show("��������� �������� �� ���������� ������������ ����������: " + Name_select, "��������� ������������", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        app = new Excel.Application();
                        Excel.Workbook xlWB = app.Workbooks.Open(desktopPath + @"\docs\Service.xls",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
                        Excel.Worksheet xlSht = xlWB.Worksheets["����1"];

                        // �������� ���������
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            connection.Open();
                            var query = string.Format(SQL_query.Comps_info, Name_select);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var cell = ds.Tables[0].Rows[0];
                            xlSht.Cells[4, 2] = cell[2].ToString();  //��� ����������
                            DateTime dt = (DateTime)cell[6];
                            xlSht.Cells[5, 2] = dt.ToShortDateString(); //���� ����� � ����
                            xlSht.Cells[3, 4] = cell[5].ToString();  //�������
                            xlSht.Cells[4, 4] = cell[4].ToString();  //ip �����
                            xlSht.Cells[5, 4] = cell[3].ToString();  //���������� �����

                            query = string.Format(SQL_query.sql_where, cell[1]);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var div = ds.Tables[0].Rows[0];
                            xlSht.Cells[3, 2] = div[1].ToString();
                            
                            query = string.Format(SQL_query.service_info, Name_select);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            int i = 0;
                            foreach (DataRow cel in ds.Tables[0].Rows)
                            {
                                dt = (DateTime)cel[2];
                                xlSht.Cells[9+i, 1] = dt.ToLongDateString();  //����
                                xlSht.Cells[9+i, 2] = cel[1].ToString();  //���
                                xlSht.Cells[9+i, 3] = cel[3].ToString();  //����������
                                i++;
                            }
                            Excel.Range range2 = xlSht.get_Range("A9", "E" + (9 + i - 1).ToString());
                            range2.EntireColumn.AutoFit();
                            range2 = xlSht.get_Range("A9", "C" + (9 + i - 1).ToString());
                            range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                        if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                            return;
                        // �������� ��������� ����
                        string filename = saveFileDialog1.FileName;
                        app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlSht = null;
                        xlWB = null;
                        app.Visible = true;
                    }
                    finally
                    {
                        

                    }
                }
            }
            else
            {
                MessageBox.Show("�������� ����������");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (ID_select > 0)
            {
                DialogResult dialogResult = MessageBox.Show("��������� �������� �� �������������� �� ����������: " + Name_select, "������������� ��", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        app = new Excel.Application();
                        Excel.Workbook xlWB = app.Workbooks.Open(desktopPath + @"\docs\Soft.xls",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
                        Excel.Worksheet xlSht = xlWB.Worksheets["����1"];

                        // �������� ���������
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            connection.Open();
                            var query = string.Format(SQL_query.Comps_info, Name_select);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var cell = ds.Tables[0].Rows[0];
                            xlSht.Cells[4, 2] = cell[2].ToString();  //��� ����������
                            DateTime dt = (DateTime)cell[6];
                            xlSht.Cells[5, 2] = dt.ToShortDateString(); //���� ����� � ����
                            xlSht.Cells[3, 4] = cell[5].ToString();  //�������
                            xlSht.Cells[4, 4] = cell[4].ToString();  //ip �����
                            xlSht.Cells[5, 4] = cell[3].ToString();  //���������� �����

                            query = string.Format(SQL_query.sql_where, cell[1]);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var div = ds.Tables[0].Rows[0];
                            xlSht.Cells[3, 2] = div[1].ToString();

                            query = string.Format(SQL_query.soft_info, Name_select);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            int i = 0;
                            foreach (DataRow cel in ds.Tables[0].Rows)
                            {

                                xlSht.Cells[9 + i, 1] = cel[1].ToString();   //��������
                                xlSht.Cells[9 + i, 2] = cel[2].ToString();  //�����-�������������
                                xlSht.Cells[9 + i, 3] = cel[3].ToString();  //����������
                                i++;
                            }
                            Excel.Range range2 = xlSht.get_Range("A9", "E" + (9 + i - 1).ToString());
                            range2.EntireColumn.AutoFit();
                            range2 = xlSht.get_Range("A9", "C" + (9 + i - 1).ToString());
                            range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                        if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                            return;
                        // �������� ��������� ����
                        string filename = saveFileDialog1.FileName;
                        app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlSht = null;
                        xlWB = null;
                        app.Visible = true;
                    }
                    finally
                    {
                        

                    }
                }
            }
            else
            {
                MessageBox.Show("�������� ����������");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (ID_div>0)
            {
                DialogResult dialogResult = MessageBox.Show("��������� �������� �� ����������� ������������� (������): " + Name_div, "�������� �� �������������� (�������)", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        app = new Excel.Application();
                        Excel.Workbook xlWB = app.Workbooks.Open(desktopPath + @"\docs\Div.xls",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
                        Excel.Worksheet xlSht = xlWB.Worksheets["����1"];

                        // �������� ���������
                        using (SqlConnection connection = new SqlConnection(SQL_query.connectionString))
                        {
                            connection.Open();
                            var query = string.Format(SQL_query.sql_div, ID_div);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            var cell = ds.Tables[0].Rows[0];
                            xlSht.Cells[1, 2] = cell[1].ToString();  //�������� �������������
                            xlSht.Cells[2, 2] = cell[2].ToString();  //���������

                            query = string.Format(SQL_query.sql_div_comp, cell[0]);
                            adapter = new SqlDataAdapter(query, connection);
                            ds = new DataSet();
                            adapter.Fill(ds);
                            int i = 0;
                            foreach (DataRow cel in ds.Tables[0].Rows)
                            {

                                xlSht.Cells[6 + i, 1] = cel[0].ToString();   //��� ��
                                xlSht.Cells[6 + i, 2] = cel[1].ToString();  //ip �����
                                xlSht.Cells[6 + i, 3] = cel[2].ToString();  //�����������
                                i++;
                            }
                            Excel.Range range2 = xlSht.get_Range("A6", "C" + (6 + i - 1).ToString());
                           range2.EntireColumn.AutoFit();
                            range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                        if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                            return;
                        // �������� ��������� ����
                        string filename = saveFileDialog1.FileName;
                        app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlSht = null;
                        xlWB = null;
                        app.Visible = true;
                    }
                    finally
                    {
                                          
                    }
                }
            }
            else
            {
                MessageBox.Show("�������� �������������(�����)");
            }
        }
    }

    class Data
    {
        public string Name { set; get; }
        public int id { set; get; }
        public Data(int id, string Name)
        {
            this.Name = Name;
            this.id = id;
        }
    }




}