namespace AnHuiSetPS
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            this.btnSetFile = new System.Windows.Forms.Button();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnSetPS = new System.Windows.Forms.Button();
            this.btnOpenWebSite = new System.Windows.Forms.Button();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnLogin = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnStop = new System.Windows.Forms.Button();
            this.dataGridView_ShowLine = new AnHuiSetPS.DataGridView_ShowLine(this.components);
            this.cbbUsers = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSetFile
            // 
            this.btnSetFile.Location = new System.Drawing.Point(286, 50);
            this.btnSetFile.Name = "btnSetFile";
            this.btnSetFile.Size = new System.Drawing.Size(109, 23);
            this.btnSetFile.TabIndex = 45;
            this.btnSetFile.Text = "选择文件";
            this.btnSetFile.UseVisualStyleBackColor = true;
            this.btnSetFile.Click += new System.EventHandler(this.btnSetFile_Click);
            // 
            // tbPath
            // 
            this.tbPath.Location = new System.Drawing.Point(16, 51);
            this.tbPath.Name = "tbPath";
            this.tbPath.ReadOnly = true;
            this.tbPath.Size = new System.Drawing.Size(256, 20);
            this.tbPath.TabIndex = 44;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.AllowUserToResizeRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(12, 99);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dataGridView_ShowLine.SetShowLineNo(this.dgv, true);
            this.dgv.Size = new System.Drawing.Size(697, 552);
            this.dgv.TabIndex = 43;
            // 
            // btnQuery
            // 
            this.btnQuery.Location = new System.Drawing.Point(859, 30);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(75, 23);
            this.btnQuery.TabIndex = 42;
            this.btnQuery.Text = "查询";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Visible = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnSetPS
            // 
            this.btnSetPS.Location = new System.Drawing.Point(496, 50);
            this.btnSetPS.Name = "btnSetPS";
            this.btnSetPS.Size = new System.Drawing.Size(75, 23);
            this.btnSetPS.TabIndex = 41;
            this.btnSetPS.Text = "设置配送";
            this.btnSetPS.UseVisualStyleBackColor = true;
            this.btnSetPS.Click += new System.EventHandler(this.btnSetPS_Click);
            // 
            // btnOpenWebSite
            // 
            this.btnOpenWebSite.Location = new System.Drawing.Point(406, 50);
            this.btnOpenWebSite.Name = "btnOpenWebSite";
            this.btnOpenWebSite.Size = new System.Drawing.Size(79, 23);
            this.btnOpenWebSite.TabIndex = 40;
            this.btnOpenWebSite.Text = "在网页中打开";
            this.btnOpenWebSite.UseVisualStyleBackColor = true;
            this.btnOpenWebSite.Click += new System.EventHandler(this.btnOpenWebSite_Click);
            // 
            // txtCode
            // 
            this.txtCode.Location = new System.Drawing.Point(425, 14);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(55, 20);
            this.txtCode.TabIndex = 37;
            this.txtCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCode_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(368, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 38;
            this.label3.Text = "验证码:";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(295, 14);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.ReadOnly = true;
            this.txtPassword.Size = new System.Drawing.Size(62, 20);
            this.txtPassword.TabIndex = 35;
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(155, 14);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.ReadOnly = true;
            this.txtUserName.Size = new System.Drawing.Size(84, 20);
            this.txtUserName.TabIndex = 33;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(250, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 36;
            this.label2.Text = "密码:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(98, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 34;
            this.label1.Text = "用户名:";
            // 
            // btnLogin
            // 
            this.btnLogin.Location = new System.Drawing.Point(577, 13);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(75, 23);
            this.btnLogin.TabIndex = 32;
            this.btnLogin.Text = "登陆";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(491, 14);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(79, 20);
            this.pictureBox1.TabIndex = 31;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // listBox1
            // 
            this.listBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(715, 97);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(256, 550);
            this.listBox1.TabIndex = 46;
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(577, 50);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(75, 23);
            this.btnStop.TabIndex = 47;
            this.btnStop.Text = "停止配送";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // cbbUsers
            // 
            this.cbbUsers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbbUsers.FormattingEnabled = true;
            this.cbbUsers.Items.AddRange(new object[] {
            "请选择",
            "康力",
            "恒天"});
            this.cbbUsers.Location = new System.Drawing.Point(16, 14);
            this.cbbUsers.Name = "cbbUsers";
            this.cbbUsers.Size = new System.Drawing.Size(76, 21);
            this.cbbUsers.TabIndex = 48;
            this.cbbUsers.SelectedIndexChanged += new System.EventHandler(this.cbbUsers_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(859, 59);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 49;
            this.button1.Text = "测试";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(975, 663);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cbbUsers);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.btnSetFile);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.btnSetPS);
            this.Controls.Add(this.btnOpenWebSite);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUserName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.pictureBox1);
            this.Name = "MainForm";
            this.Text = "安徽设置配送";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSetFile;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnSetPS;
        private System.Windows.Forms.Button btnOpenWebSite;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ListBox listBox1;
        private DataGridView_ShowLine dataGridView_ShowLine;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.ComboBox cbbUsers;
        private System.Windows.Forms.Button button1;
    }
}

