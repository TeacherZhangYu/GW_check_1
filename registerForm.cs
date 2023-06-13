using KJ_chenk;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace GW_check
{
    public partial class registerForm : Form
    {
        private  Xml xmlfile = new Xml();
        public registerForm()
        {
            InitializeComponent();
        }
        private void register_exsit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btn_register_ture_Click(object sender, EventArgs e)
        {
           
            if(registerProject.Text!="")
            {
                if (registerPassword.Text != "")
                {
                    if (register_project.Text != "")
                    {
                        try
                        {
                            xmlfile.AppendNode(registerProject.Text, registerPassword.Text, register_project.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                            MessageBox.Show("新用户注册成功！！");
                    }
                    else
                    {
                        MessageBox.Show("项目不能为空！！");
                    }
                }
                else
                {
                    MessageBox.Show("密码不能为空！！");
                }
            }
            else
            {
                MessageBox.Show("单位不能为空！！");
            }                                             
        }
    }
}
