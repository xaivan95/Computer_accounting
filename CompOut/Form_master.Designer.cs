namespace CompOut
{
    partial class Form_master
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label120 = new System.Windows.Forms.Label();
            this.textBox85 = new System.Windows.Forms.TextBox();
            this.label122 = new System.Windows.Forms.Label();
            this.button15 = new System.Windows.Forms.Button();
            this.button16 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(503, 350);
            this.dataGridView1.TabIndex = 0;
            // 
            // label120
            // 
            this.label120.AutoSize = true;
            this.label120.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label120.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label120.Location = new System.Drawing.Point(350, 364);
            this.label120.Name = "label120";
            this.label120.Size = new System.Drawing.Size(78, 15);
            this.label120.TabIndex = 24;
            this.label120.Text = "Примечание";
            // 
            // textBox85
            // 
            this.textBox85.Location = new System.Drawing.Point(257, 382);
            this.textBox85.Name = "textBox85";
            this.textBox85.Size = new System.Drawing.Size(246, 23);
            this.textBox85.TabIndex = 23;
            // 
            // label122
            // 
            this.label122.AutoSize = true;
            this.label122.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label122.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label122.Location = new System.Drawing.Point(81, 364);
            this.label122.Name = "label122";
            this.label122.Size = new System.Drawing.Size(92, 15);
            this.label122.TabIndex = 22;
            this.label122.Text = "Ф.И.О. мастера";
            // 
            // button15
            // 
            this.button15.Image = global::CompOut.Properties.Resources.delete_1;
            this.button15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button15.Location = new System.Drawing.Point(330, 411);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(135, 26);
            this.button15.TabIndex = 21;
            this.button15.Text = "Удалить";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Click += new System.EventHandler(this.button15_Click);
            // 
            // button16
            // 
            this.button16.Image = global::CompOut.Properties.Resources.plus1;
            this.button16.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button16.Location = new System.Drawing.Point(57, 411);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(135, 26);
            this.button16.TabIndex = 20;
            this.button16.Text = "Добавить";
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Click += new System.EventHandler(this.button16_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(0, 382);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(251, 23);
            this.textBox1.TabIndex = 25;
            // 
            // Form_master
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(503, 462);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label120);
            this.Controls.Add(this.textBox85);
            this.Controls.Add(this.label122);
            this.Controls.Add(this.button15);
            this.Controls.Add(this.button16);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form_master";
            this.Text = "Мастера";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form_master_FormClosed);
            this.Shown += new System.EventHandler(this.Form_master_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataGridView dataGridView1;
        private Label label120;
        private TextBox textBox85;
        private Label label122;
        private Button button15;
        private Button button16;
        private TextBox textBox1;
    }
}