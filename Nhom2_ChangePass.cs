//Nhóm 2: Nguyễn Hoàng Minh: MSV - TinA6
//Cao Trung Hiếu: MSV - TinA6
//Bùi Tuấn Thành: MSV - TinA6

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace ThiTN
{
    public partial class Nhom2_ChangePass : Form
    {
        static SqlConnection con = new SqlConnection(@"Data Source=LAPTOP-AOALA04D\SEPTSQL;Initial Catalog=qlsv_2 _NHMinhCTHieuBTThanh_LTNET;Integrated Security=True");
        public Nhom2_ChangePass()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (t1.Text == "" && t2.Text == "")
            {
                MessageBox.Show("Hãy nhập tài khoản và mật khẩu!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (t1.Text == "")
            {
                MessageBox.Show("Bạn chưa nhập tên tài khoản!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (t2.Text == "")
            {
                MessageBox.Show("Bạn chưa nhập mật khẩu!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("Select * from login where username=@username and password=@password", con);
                    cmd.Parameters.AddWithValue("@username", t1.Text);
                    cmd.Parameters.AddWithValue("@password", t2.Text);
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    DataTable ds = new DataTable();
                    adapt.Fill(ds);
                    if (ds.Rows.Count.ToString() == "1")
                    {
                        if (t3.Text == t4.Text)
                        {
                            con.Open();
                            SqlCommand cmdx = new SqlCommand("UPDATE login SET password='" + t4.Text + "' where password = '" + t1.Text + "'", con);

                            cmdx.ExecuteNonQuery();
                            con.Close();
                            MessageBox.Show("Đổi mật khẩu thành công!");
                        }
                        else
                        {
                            MessageBox.Show("Mật khẩu mới và Nhập lại mật khẩu phải giống nhau");
                        }

                    }
                    else
                    {
                        MessageBox.Show("Bạn chưa nhập đúng thông tin tài khoản!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }
                catch (Exception f)
                {
                    MessageBox.Show(f.Message);
                }
            }
            
        }
    }
}
