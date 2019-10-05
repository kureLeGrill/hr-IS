namespace LEGAL
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.контрактыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сделатьКонтрактToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.фактурыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сделатьФактуруToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.фирмаCZToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.профToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(13, 577);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(190, 58);
            this.button1.TabIndex = 0;
            this.button1.Text = "ЛЕГАЛИЗАЦИЯ";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(241, 611);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(154, 24);
            this.checkBox1.TabIndex = 1;
            this.checkBox1.Text = "Только цестяки";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.контрактыToolStripMenuItem,
            this.фактурыToolStripMenuItem,
            this.добавитьToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(9, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(463, 35);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // контрактыToolStripMenuItem
            // 
            this.контрактыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.сделатьКонтрактToolStripMenuItem});
            this.контрактыToolStripMenuItem.Name = "контрактыToolStripMenuItem";
            this.контрактыToolStripMenuItem.Size = new System.Drawing.Size(111, 29);
            this.контрактыToolStripMenuItem.Text = "Контракты";
            // 
            // сделатьКонтрактToolStripMenuItem
            // 
            this.сделатьКонтрактToolStripMenuItem.Name = "сделатьКонтрактToolStripMenuItem";
            this.сделатьКонтрактToolStripMenuItem.Size = new System.Drawing.Size(238, 30);
            this.сделатьКонтрактToolStripMenuItem.Text = "Сделать контракт";
            this.сделатьКонтрактToolStripMenuItem.Click += new System.EventHandler(this.сделатьКонтрактToolStripMenuItem_Click);
            // 
            // фактурыToolStripMenuItem
            // 
            this.фактурыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.сделатьФактуруToolStripMenuItem});
            this.фактурыToolStripMenuItem.Name = "фактурыToolStripMenuItem";
            this.фактурыToolStripMenuItem.Size = new System.Drawing.Size(95, 29);
            this.фактурыToolStripMenuItem.Text = "Фактуры";
            // 
            // сделатьФактуруToolStripMenuItem
            // 
            this.сделатьФактуруToolStripMenuItem.Name = "сделатьФактуруToolStripMenuItem";
            this.сделатьФактуруToolStripMenuItem.Size = new System.Drawing.Size(231, 30);
            this.сделатьФактуруToolStripMenuItem.Text = "Сделать фактуру";
            this.сделатьФактуруToolStripMenuItem.Click += new System.EventHandler(this.сделатьФактуруToolStripMenuItem_Click);
            // 
            // добавитьToolStripMenuItem
            // 
            this.добавитьToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.фирмаCZToolStripMenuItem,
            this.профToolStripMenuItem});
            this.добавитьToolStripMenuItem.Name = "добавитьToolStripMenuItem";
            this.добавитьToolStripMenuItem.Size = new System.Drawing.Size(102, 29);
            this.добавитьToolStripMenuItem.Text = "Добавить";
            // 
            // фирмаCZToolStripMenuItem
            // 
            this.фирмаCZToolStripMenuItem.Name = "фирмаCZToolStripMenuItem";
            this.фирмаCZToolStripMenuItem.Size = new System.Drawing.Size(252, 30);
            this.фирмаCZToolStripMenuItem.Text = "Фирма партнер";
            this.фирмаCZToolStripMenuItem.Click += new System.EventHandler(this.фирмаCZToolStripMenuItem_Click);
            // 
            // профToolStripMenuItem
            // 
            this.профToolStripMenuItem.Name = "профToolStripMenuItem";
            this.профToolStripMenuItem.Size = new System.Drawing.Size(252, 30);
            this.профToolStripMenuItem.Text = "Профессия";
            this.профToolStripMenuItem.Click += new System.EventHandler(this.профToolStripMenuItem_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
            this.dataGridView1.Location = new System.Drawing.Point(13, 79);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(438, 480);
            this.dataGridView1.TabIndex = 3;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Профессия";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Зарплата";
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Часы работы";
            this.Column3.Name = "Column3";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(463, 649);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.Text = "LEHA-5000";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem контрактыToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сделатьКонтрактToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem фактурыToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сделатьФактуруToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.ToolStripMenuItem добавитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem фирмаCZToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem профToolStripMenuItem;
    }
}

