
namespace GW_check
{
    partial class registerForm
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_register_ture = new System.Windows.Forms.Button();
            this.registerProject = new System.Windows.Forms.TextBox();
            this.registerPassword = new System.Windows.Forms.TextBox();
            this.register_project = new System.Windows.Forms.TextBox();
            this.register_exsit = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(128)))), ((int)(((byte)(158)))));
            this.panel1.Controls.Add(this.register_exsit);
            this.panel1.Controls.Add(this.register_project);
            this.panel1.Controls.Add(this.registerPassword);
            this.panel1.Controls.Add(this.registerProject);
            this.panel1.Controls.Add(this.btn_register_ture);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(-2, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(526, 305);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(154, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(213, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户信息注册";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(107, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 19);
            this.label2.TabIndex = 1;
            this.label2.Text = "单位：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(107, 128);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(69, 19);
            this.label3.TabIndex = 2;
            this.label3.Text = "密码：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(107, 172);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 19);
            this.label4.TabIndex = 3;
            this.label4.Text = "项目：";
            // 
            // btn_register_ture
            // 
            this.btn_register_ture.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_register_ture.Location = new System.Drawing.Point(144, 234);
            this.btn_register_ture.Name = "btn_register_ture";
            this.btn_register_ture.Size = new System.Drawing.Size(85, 46);
            this.btn_register_ture.TabIndex = 4;
            this.btn_register_ture.Text = "确定";
            this.btn_register_ture.UseVisualStyleBackColor = true;
            this.btn_register_ture.Click += new System.EventHandler(this.btn_register_ture_Click);
            // 
            // registerProject
            // 
            this.registerProject.Location = new System.Drawing.Point(185, 84);
            this.registerProject.Name = "registerProject";
            this.registerProject.Size = new System.Drawing.Size(191, 21);
            this.registerProject.TabIndex = 5;
            // 
            // registerPassword
            // 
            this.registerPassword.Location = new System.Drawing.Point(185, 126);
            this.registerPassword.Name = "registerPassword";
            this.registerPassword.Size = new System.Drawing.Size(191, 21);
            this.registerPassword.TabIndex = 6;
            // 
            // register_project
            // 
            this.register_project.Location = new System.Drawing.Point(185, 170);
            this.register_project.Name = "register_project";
            this.register_project.Size = new System.Drawing.Size(191, 21);
            this.register_project.TabIndex = 7;
            // 
            // register_exsit
            // 
            this.register_exsit.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.register_exsit.Location = new System.Drawing.Point(276, 234);
            this.register_exsit.Name = "register_exsit";
            this.register_exsit.Size = new System.Drawing.Size(85, 46);
            this.register_exsit.TabIndex = 8;
            this.register_exsit.Text = "退出";
            this.register_exsit.UseVisualStyleBackColor = true;
            this.register_exsit.Click += new System.EventHandler(this.register_exsit_Click);
            // 
            // registerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(523, 303);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "registerForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "registerForm";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button register_exsit;
        private System.Windows.Forms.TextBox register_project;
        private System.Windows.Forms.TextBox registerPassword;
        private System.Windows.Forms.TextBox registerProject;
        private System.Windows.Forms.Button btn_register_ture;
    }
}