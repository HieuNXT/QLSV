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
using System.Diagnostics;

namespace ThiTN
{
    public partial class Nhom2_Login : Form
    {
        static SqlConnection con = new SqlConnection(@"Data Source=LAPTOP-AOALA04D\SEPTSQL;Initial Catalog=qlsv_2 _NHMinhCTHieuBTThanh_LTNET;Integrated Security=True");
        public Nhom2_Login()
        {
            InitializeComponent();
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text == "" && txtPassword.Text == "")
            {
                MessageBox.Show("Hãy nhập tài khoản và mật khẩu!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (txtUsername.Text == "")
            {
                MessageBox.Show("Bạn chưa nhập tên tài khoản!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (txtPassword.Text == "")
            {
                MessageBox.Show("Bạn chưa nhập mật khẩu!", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SqlCommand cmd = new SqlCommand("Select * from login where username=@username and password=@password", con);
                cmd.Parameters.AddWithValue("@username", txtUsername.Text);
                cmd.Parameters.AddWithValue("@password", txtPassword.Text);
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                DataTable ds = new DataTable();
                adapt.Fill(ds);

                if (ds.Rows.Count == 1)
                {
                    string a;
                    a = ds.Rows[0]["state"].ToString();
                    if (a == "Mở")
                    {
                        switch (ds.Rows[0]["mod"] as string)
                        {
                            case "Admin":
                                {
                                    MessageBox.Show("Đăng nhập thành công!");
                                    Nhom2_MainUI ss = new Nhom2_MainUI();
                                    ss.Show();
                                    ss.label13.Text = Convert.ToString(ds.Rows[0]["fullname"]);
                                    ss.label15.Text = Convert.ToString(ds.Rows[0]["fullname"]);
                                    this.Hide();
                                    break;
                                }

                            case "User":
                                {

                                    MessageBox.Show("Đăng nhập thành công!");
                                    Nhom2_MainUI ss = new Nhom2_MainUI();
                                    ss.tool2.Visible = false;
                                    ss.label13.Text = Convert.ToString(ds.Rows[0]["fullname"]);
                                    ss.label15.Text = Convert.ToString(ds.Rows[0]["fullname"]);
                                    ss.Show();
                                    this.Hide();
                                    break;

                                }

                            default:
                                {
                                    MessageBox.Show("Không thể xác minh tài khoản!",
                                                 "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Tài khoản của bạn đã bị khóa!",
                      "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
