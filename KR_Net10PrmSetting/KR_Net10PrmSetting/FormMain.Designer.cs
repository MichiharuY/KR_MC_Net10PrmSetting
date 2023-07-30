namespace KR_Net10PrmSetting
{
    partial class FormMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txtAddress1 = new TextBox();
            cmbItemName1 = new ComboBox();
            lblParamName = new Label();
            lblValue = new Label();
            btnSet1 = new Button();
            txtSetValue1 = new TextBox();
            label1 = new Label();
            txtAddress2 = new TextBox();
            txtSetValue2 = new TextBox();
            cmbItemName2 = new ComboBox();
            label3 = new Label();
            label5 = new Label();
            label6 = new Label();
            btnSet2 = new Button();
            groupBox1 = new GroupBox();
            txtNet10Value1 = new TextBox();
            label2 = new Label();
            groupBox2 = new GroupBox();
            txtNet10Value2 = new TextBox();
            label4 = new Label();
            btnExcelReload = new Button();
            lblStatus = new Label();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // txtAddress1
            // 
            txtAddress1.Location = new Point(243, 49);
            txtAddress1.Margin = new Padding(2);
            txtAddress1.Name = "txtAddress1";
            txtAddress1.ReadOnly = true;
            txtAddress1.Size = new Size(73, 23);
            txtAddress1.TabIndex = 0;
            txtAddress1.TextAlign = HorizontalAlignment.Center;
            // 
            // cmbItemName1
            // 
            cmbItemName1.FormattingEnabled = true;
            cmbItemName1.Location = new Point(13, 49);
            cmbItemName1.Margin = new Padding(2);
            cmbItemName1.Name = "cmbItemName1";
            cmbItemName1.Size = new Size(227, 23);
            cmbItemName1.TabIndex = 1;
            cmbItemName1.SelectedIndexChanged += cmbItemName1_SelectedIndexChanged;
            // 
            // lblParamName
            // 
            lblParamName.AutoSize = true;
            lblParamName.Location = new Point(87, 24);
            lblParamName.Margin = new Padding(2, 0, 2, 0);
            lblParamName.Name = "lblParamName";
            lblParamName.Size = new Size(43, 15);
            lblParamName.TabIndex = 2;
            lblParamName.Text = "項目名";
            // 
            // lblValue
            // 
            lblValue.AutoSize = true;
            lblValue.Location = new Point(243, 24);
            lblValue.Margin = new Padding(2, 0, 2, 0);
            lblValue.Name = "lblValue";
            lblValue.Size = new Size(75, 15);
            lblValue.TabIndex = 2;
            lblValue.Text = "NET10アドレス";
            // 
            // btnSet1
            // 
            btnSet1.BackColor = Color.Teal;
            btnSet1.ForeColor = Color.White;
            btnSet1.Location = new Point(542, 48);
            btnSet1.Margin = new Padding(2);
            btnSet1.Name = "btnSet1";
            btnSet1.Size = new Size(78, 25);
            btnSet1.TabIndex = 3;
            btnSet1.Text = "設定";
            btnSet1.UseVisualStyleBackColor = false;
            btnSet1.Click += btnSet1_Click;
            // 
            // txtSetValue1
            // 
            txtSetValue1.Location = new Point(428, 49);
            txtSetValue1.Margin = new Padding(2);
            txtSetValue1.Name = "txtSetValue1";
            txtSetValue1.Size = new Size(106, 23);
            txtSetValue1.TabIndex = 0;
            txtSetValue1.TextAlign = HorizontalAlignment.Right;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(460, 24);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(43, 15);
            label1.TabIndex = 2;
            label1.Text = "設定値";
            // 
            // txtAddress2
            // 
            txtAddress2.Location = new Point(243, 45);
            txtAddress2.Margin = new Padding(2);
            txtAddress2.Name = "txtAddress2";
            txtAddress2.ReadOnly = true;
            txtAddress2.Size = new Size(73, 23);
            txtAddress2.TabIndex = 0;
            txtAddress2.TextAlign = HorizontalAlignment.Center;
            // 
            // txtSetValue2
            // 
            txtSetValue2.Location = new Point(428, 45);
            txtSetValue2.Margin = new Padding(2);
            txtSetValue2.Name = "txtSetValue2";
            txtSetValue2.Size = new Size(106, 23);
            txtSetValue2.TabIndex = 0;
            txtSetValue2.TextAlign = HorizontalAlignment.Right;
            // 
            // cmbItemName2
            // 
            cmbItemName2.FormattingEnabled = true;
            cmbItemName2.Location = new Point(13, 45);
            cmbItemName2.Margin = new Padding(2);
            cmbItemName2.Name = "cmbItemName2";
            cmbItemName2.Size = new Size(227, 23);
            cmbItemName2.TabIndex = 1;
            cmbItemName2.SelectedIndexChanged += cmbItemName2_SelectedIndexChanged;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(87, 20);
            label3.Margin = new Padding(2, 0, 2, 0);
            label3.Name = "label3";
            label3.Size = new Size(43, 15);
            label3.TabIndex = 2;
            label3.Text = "項目名";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(243, 20);
            label5.Margin = new Padding(2, 0, 2, 0);
            label5.Name = "label5";
            label5.Size = new Size(75, 15);
            label5.TabIndex = 2;
            label5.Text = "NET10アドレス";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(460, 20);
            label6.Margin = new Padding(2, 0, 2, 0);
            label6.Name = "label6";
            label6.Size = new Size(43, 15);
            label6.TabIndex = 2;
            label6.Text = "設定値";
            // 
            // btnSet2
            // 
            btnSet2.BackColor = Color.Teal;
            btnSet2.ForeColor = Color.White;
            btnSet2.Location = new Point(542, 44);
            btnSet2.Margin = new Padding(2);
            btnSet2.Name = "btnSet2";
            btnSet2.Size = new Size(78, 25);
            btnSet2.TabIndex = 3;
            btnSet2.Text = "設定";
            btnSet2.UseVisualStyleBackColor = false;
            btnSet2.Click += btnSet2_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblParamName);
            groupBox1.Controls.Add(txtNet10Value1);
            groupBox1.Controls.Add(txtAddress1);
            groupBox1.Controls.Add(btnSet1);
            groupBox1.Controls.Add(txtSetValue1);
            groupBox1.Controls.Add(cmbItemName1);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(lblValue);
            groupBox1.Location = new Point(21, 17);
            groupBox1.Margin = new Padding(2);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(2);
            groupBox1.Size = new Size(632, 90);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            groupBox1.Text = "制御パラメータ";
            // 
            // txtNet10Value1
            // 
            txtNet10Value1.Location = new Point(319, 49);
            txtNet10Value1.Margin = new Padding(2);
            txtNet10Value1.Name = "txtNet10Value1";
            txtNet10Value1.ReadOnly = true;
            txtNet10Value1.Size = new Size(106, 23);
            txtNet10Value1.TabIndex = 0;
            txtNet10Value1.TextAlign = HorizontalAlignment.Right;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(327, 24);
            label2.Margin = new Padding(2, 0, 2, 0);
            label2.Name = "label2";
            label2.Size = new Size(89, 15);
            label2.TabIndex = 2;
            label2.Text = "NET10デバイス値";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(label3);
            groupBox2.Controls.Add(txtNet10Value2);
            groupBox2.Controls.Add(txtAddress2);
            groupBox2.Controls.Add(btnSet2);
            groupBox2.Controls.Add(txtSetValue2);
            groupBox2.Controls.Add(label4);
            groupBox2.Controls.Add(label6);
            groupBox2.Controls.Add(cmbItemName2);
            groupBox2.Controls.Add(label5);
            groupBox2.Location = new Point(21, 121);
            groupBox2.Margin = new Padding(2);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new Padding(2);
            groupBox2.Size = new Size(632, 90);
            groupBox2.TabIndex = 5;
            groupBox2.TabStop = false;
            groupBox2.Text = "警報パラメータ";
            // 
            // txtNet10Value2
            // 
            txtNet10Value2.Location = new Point(319, 45);
            txtNet10Value2.Margin = new Padding(2);
            txtNet10Value2.Name = "txtNet10Value2";
            txtNet10Value2.ReadOnly = true;
            txtNet10Value2.Size = new Size(106, 23);
            txtNet10Value2.TabIndex = 0;
            txtNet10Value2.TextAlign = HorizontalAlignment.Right;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(328, 20);
            label4.Margin = new Padding(2, 0, 2, 0);
            label4.Name = "label4";
            label4.Size = new Size(89, 15);
            label4.TabIndex = 2;
            label4.Text = "NET10デバイス値";
            // 
            // btnExcelReload
            // 
            btnExcelReload.Location = new Point(558, 223);
            btnExcelReload.Name = "btnExcelReload";
            btnExcelReload.Size = new Size(95, 23);
            btnExcelReload.TabIndex = 6;
            btnExcelReload.Text = "EXCEL再読込";
            btnExcelReload.UseVisualStyleBackColor = true;
            btnExcelReload.Click += btnExcelReload_Click;
            // 
            // lblStatus
            // 
            lblStatus.Location = new Point(24, 228);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(175, 13);
            lblStatus.TabIndex = 8;
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(670, 260);
            Controls.Add(lblStatus);
            Controls.Add(btnExcelReload);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Margin = new Padding(2);
            Name = "FormMain";
            Text = "NET10パラメータ設定";
            Activated += FormMain_Activated;
            FormClosing += FormMain_FormClosing;
            Load += FormMain_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TextBox txtAddress1;
        private ComboBox cmbItemName1;
        private Label lblParamName;
        private Label lblValue;
        private Button btnSet1;
        private TextBox txtSetValue1;
        private Label label1;
        private TextBox txtAddress2;
        private TextBox txtSetValue2;
        private ComboBox cmbItemName2;
        private Label label3;
        private Label label5;
        private Label label6;
        private Button btnSet2;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private TextBox txtNet10Value1;
        private Label label2;
        private TextBox txtNet10Value2;
        private Label label4;
        private Button btnExcelReload;
        private Label lblStatus;
    }
}