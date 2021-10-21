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
using System.Globalization;
using SautinSoft.Document;


namespace ThiTN
{
    public partial class Nhom2_MainUI : Form
    {
        static SqlConnection con = new SqlConnection(@"Data Source=LAPTOP-AOALA04D\SEPTSQL;Initial Catalog=qlsv_2 _NHMinhCTHieuBTThanh_LTNET;Integrated Security=True");
        int row;
        public Nhom2_MainUI()
        {
            InitializeComponent();
        }

        private void MainUI_Load(object sender, EventArgs e)
        {
            loadview();
            loadview1();
            loadview2();
        }
        private void loadview()
        {
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM login", con);
            DataTable dtb = new DataTable();
            da.Fill(dtb);
            dataGridView1.DataSource = dtb;
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            con.Open();
            if (button1.Text == "Thêm")
            {
                string qur = "INSERT INTO login VALUES ('" + t1.Text + "','" + t2.Text + "',N'" + t3.Text + "','" + t4.Text + "','" + c1.Text + "',N'" + c2.Text + "')";

                SqlCommand cmd = new SqlCommand(qur, con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Thêm thành công");
            }
            else
            {
                string up = "UPDATE login SET username=@username, password=@password, fullname=@fullname, mail=@mail, mod=@mod, state=@state   WHERE username=@username";
                SqlCommand cmd = new SqlCommand(up, con);
                cmd.Parameters.AddWithValue("@username", t1.Text);
                cmd.Parameters.AddWithValue("@password", t2.Text);
                cmd.Parameters.AddWithValue("@fullname", t3.Text);
                cmd.Parameters.AddWithValue("@mod", c1.Text);
                cmd.Parameters.AddWithValue("@mail", t4.Text);
                cmd.Parameters.AddWithValue("@state", c2.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Sửa thành công");
            }
            loadview();
        }
             private void button2_Click_1(object sender, EventArgs e)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("DELETE FROM login WHERE username=@username ", con);
                cmd.Parameters.AddWithValue("@username", t1.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                loadview();
            }
        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {
            t1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            t2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            t3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            t4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            c1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();

            button1.Text = "Sửa";
            button2.Enabled = true;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            t1.Clear();
            t2.Clear();
            t3.Clear();
            t4.Clear();
            c1.Text = "Quyền";
        }

        private void thoátToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
//-----------------------------------------------------------END OF LOGIN AND ACCOUNTES MANAGEMENT-------------------------------------------------------------
        private void button11_Click(object sender, EventArgs e)
        {
            con.Open();
           
            if (button11.Text == "Thêm")
            {

                string qur = "INSERT INTO sinhvien VALUES ('" + txtMsv.Text + "',N'" + txtName.Text + "','" + txtBornDate.Text + "',N'" + cbxSex.Text + "',N'" + txtQue.Text + "','" + txtDrl.Text + "','" + comboBox1.Text + "')";
                
                
                SqlCommand cmd = new SqlCommand(qur, con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Thêm thành công");
            }
            else
            {
                string up = "UPDATE sinhvien SET masv=@masv, hoten=@hoten, ngaysinh=@ngaysinh, gioitinh=@gioitinh, quequan=@quequan, diemrl=@diemrl, malop=@malop    WHERE masv=@masv";
                SqlCommand cmd = new SqlCommand(up, con);
                cmd.Parameters.AddWithValue("@masv", txtMsv.Text);
                cmd.Parameters.AddWithValue("@hoten", txtName.Text);
                cmd.Parameters.AddWithValue("@ngaysinh", txtBornDate.Text);
                cmd.Parameters.AddWithValue("@gioitinh", cbxSex.Text);
                cmd.Parameters.AddWithValue("@quequan", txtQue.Text);
                cmd.Parameters.AddWithValue("@malop", comboBox1.Text);
                cmd.Parameters.AddWithValue("@diemrl", txtDrl.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Sửa thành công");
            }
            loadview1();
        }
        private void loadview1()
        {
            SqlDataAdapter db = new SqlDataAdapter("SELECT * FROM sinhvien", con);
            DataTable dta = new DataTable();
            db.Fill(dta);
            DSSV.DataSource = dta;
            DSSV.Columns[0].HeaderCell.Value = "Mã Sinh viên";
            DSSV.Columns[1].HeaderCell.Value = "Họ tên";
            DSSV.Columns[2].HeaderCell.Value = "Ngày sinh";
            DSSV.Columns[3].HeaderCell.Value = "Giới tính";
            DSSV.Columns[4].HeaderCell.Value = "Quê quán";
            DSSV.Columns[5].HeaderCell.Value = "Điểm rl";
            DSSV.Columns[6].HeaderCell.Value = "Mã lớp";

            
        }
        void DSSV1()
        {
            txtMsv.Text = DSSV.SelectedCells[0].Value.ToString();
            txtName.Text = DSSV.SelectedCells[1].Value.ToString();
            txtBornDate.Text = DSSV.SelectedCells[2].Value.ToString();
            cbxSex.Text = DSSV.SelectedCells[3].Value.ToString();
            txtQue.Text = DSSV.SelectedCells[4].Value.ToString();
            txtDrl.Text = DSSV.SelectedCells[5].Value.ToString();
            comboBox1.Text = DSSV.SelectedCells[6].Value.ToString();
        }
        private void DSSV_DoubleClick(object sender, EventArgs e)
        {
            row = DSSV.CurrentRow.Index;
            txtMsv.Text = DSSV.CurrentRow.Cells[0].Value.ToString();
            txtName.Text = DSSV.CurrentRow.Cells[1].Value.ToString();
            txtBornDate.Text = DSSV.CurrentRow.Cells[2].Value.ToString();
            cbxSex.Text = DSSV.CurrentRow.Cells[3].Value.ToString();
            txtQue.Text = DSSV.CurrentRow.Cells[4].Value.ToString();
            txtDrl.Text = DSSV.CurrentRow.Cells[5].Value.ToString();
            comboBox1.Text = DSSV.CurrentRow.Cells[6].Value.ToString();

            button11.Text = "Sửa";
            button10.Enabled = true;
            textBox1.Visible = false;
            textBox2.Visible = false;
        }
        private void button10_Click(object sender, EventArgs e)
        {
            con.Open();
            var confirmResult = MessageBox.Show("Bạn có chắc muốn xóa?",
                                        "Xác nhận xóa!!",
                                        MessageBoxButtons.YesNo);
            if (confirmResult == DialogResult.Yes)
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM sinhvien WHERE masv=@masv ", con);
                cmd.Parameters.AddWithValue("@masv", txtMsv.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                loadview1();
            }
            else
            {
                con.Close();
                loadview1();
            }
            

            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            txtMsv.Clear();
            txtName.Clear();
            txtBornDate.Clear();
            cbxSex.Text = null;
            txtQue.Clear();
            txtDrl.Clear();
            comboBox1.Text = null;
            textBox1.Visible = true;
            textBox2.Visible = true;
            button11.Text = "Thêm";
            loadview1();

        }

        private void cbxSex_MouseClick(object sender, MouseEventArgs e)
        {
            textBox1.Visible = false;
        }
        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            textBox2.Visible = false;
        }

        private void btnSearchInfo_Click(object sender, EventArgs e)
        {
            using (DataTable dt = new DataTable("Sinhvien"))
            {
                if (radioButton1.Checked == true)
                {
                    using (SqlCommand cmd = new SqlCommand("select *from sinhvien where hoten=@hoten", con))
                    {
                        cmd.Parameters.AddWithValue("@hoten", txtSearch.Text);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dt);
                        DSSV.DataSource = dt;

                    }
                }
                 else if (radioButton2.Checked == true)
                {
                    using (SqlCommand cmd = new SqlCommand("select *from sinhvien where quequan=@quequan", con))
                    {
                        cmd.Parameters.AddWithValue("@quequan", txtSearch.Text);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dt);
                        DSSV.DataSource = dt;
                    }
                }
                else
                {

                    MessageBox.Show("Vui lòng nhấn chọn tìm kiếm theo tên hoặc quê quán!!");
                }
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (row >= 0)
            {
                if (row != 0)
                {
                    DSSV.Rows[row].Selected = false;
                    DSSV.Rows[--row].Selected = true;
                }
                else
                {
                    MessageBox.Show("Hết danh sách!");
                }
            }
            DSSV1();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (row < DSSV.RowCount)
            {
                if (row != DSSV.RowCount - 1)
                {
                    DSSV.Rows[row].Selected = false;
                    DSSV.Rows[++row].Selected = true;
                }
                else
                {
                    MessageBox.Show("Hết danh sách!");
                }
                DSSV1();
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DSSV.Rows[0].Selected = true;
            DSSV1();
            loadview1();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            row = DSSV.CurrentRow.Index;
            int a;
            a = DSSV.RowCount;
            DSSV.Rows[a-1].Selected = true;
            DSSV1();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.DSSV.Sort(this.DSSV.Columns["diemrl"], ListSortDirection.Descending);
        }

//--------------------------------------------------------------------END OF SINHVIEN MANAGEMENT--------------------------------------------------------------------
        private void DSLOP_DoubleClick(object sender, EventArgs e)
        {
            txtmalop.Text = DSLOP.CurrentRow.Cells[0].Value.ToString();
            txtTenlop.Text = DSLOP.CurrentRow.Cells[1].Value.ToString();
            txtMaillop.Text = DSLOP.CurrentRow.Cells[2].Value.ToString();
            txtTruonglop.Text = DSLOP.CurrentRow.Cells[3].Value.ToString();
            button19.Text = "Sửa";
            button18.Enabled = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            con.Open();
            if (button1.Text == "Thêm")
            {
                string qur = "INSERT INTO lop VALUES ('" + txtmalop.Text + "','" + txtTenlop.Text + "','" + txtMaillop.Text + "','" + txtTruonglop.Text + "')";

                SqlCommand cmd = new SqlCommand(qur, con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Thêm thành công");
            }
            else
            {
                string up = "UPDATE lop SET malop=@malop, tenlop=@tenlop, email=@email, hotenlt=@hotenlt  WHERE malop=@malop";
                SqlCommand cmd = new SqlCommand(up, con);
                cmd.Parameters.AddWithValue("@malop", txtmalop.Text);
                cmd.Parameters.AddWithValue("@tenlop", txtTenlop.Text);
                cmd.Parameters.AddWithValue("@email", txtMaillop.Text);
                cmd.Parameters.AddWithValue("@hotenlt", txtTruonglop.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Sửa thành công");
            }
            loadview2();
        }
        private void loadview2()
        {
            con.Open();
            SqlDataAdapter db = new SqlDataAdapter("SELECT * FROM lop", con);
            DataTable dta = new DataTable();
            db.Fill(dta);
            DSLOP.DataSource = dta;
            DSLOP.Columns[0].HeaderCell.Value = "Mã Lớp";
            DSLOP.Columns[1].HeaderCell.Value = "Tên lớp";
            DSLOP.Columns[2].HeaderCell.Value = "Email lớp";
            DSLOP.Columns[3].HeaderCell.Value = "Tên tml lớp trưởng";
            comboBox1.DataSource = dta;
            comboBox1.DisplayMember = "malop";

            con.Close();
        }


        private void button18_Click(object sender, EventArgs e)
        {
            con.Open();
            var confirmResult = MessageBox.Show("Bạn có chắc muốn xóa?",
                                        "Xác nhận xóa!!",
                                        MessageBoxButtons.YesNo);
            if (confirmResult == DialogResult.Yes)
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM lop WHERE malop=@malop ", con);
                cmd.Parameters.AddWithValue("@malop", txtmalop.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                loadview2();
            }
            else
            {
                con.Close();
                loadview2();
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            txtTruonglop.Clear();
            txtMaillop.Clear();
            txtTenlop.Clear();
            txtmalop.Clear();
            button19.Text = "Thêm";

        }

        private void button26_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable("Sinhvien");
            SqlCommand cmd = new SqlCommand("select *from lop where tenlop=@tenlop", con);
            cmd.Parameters.AddWithValue("@quequan", txtSearch.Text);
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            adapter.Fill(dt);
            DSLOP.DataSource = dt;
                
            
        }

        private void button21_Click(object sender, EventArgs e)
        {
            this.DSLOP.Sort(this.DSLOP.Columns["tenlop"], ListSortDirection.Descending);
        }
        void DSlop1()
        {
            txtTruonglop.Text = DSLOP.SelectedCells[0].Value.ToString();
            txtMaillop.Text = DSLOP.SelectedCells[1].Value.ToString();
            txtTenlop.Text = DSLOP.SelectedCells[2].Value.ToString();
            txtmalop.Text = DSLOP.SelectedCells[3].Value.ToString();
        }
        private void button25_Click(object sender, EventArgs e)
        {
            DSLOP.Rows[0].Selected = true;
            DSlop1();
            loadview2();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (row >= 0)
            {
                if (row != 0)
                {
                    DSLOP.Rows[row].Selected = false;
                    DSLOP.Rows[--row].Selected = true;
                }
                else
                {
                    MessageBox.Show("Hết danh sách!");
                }
            }
            
                DSlop1();
            
        }

        private void button23_Click(object sender, EventArgs e)
        {
        if (row < DSLOP.RowCount)
        {
            if (row != DSLOP.RowCount - 1)
            {
                DSLOP.Rows[row].Selected = false;
                DSLOP.Rows[++row].Selected = true;
            }
            else
            {
                MessageBox.Show("Hết danh sách!");
            }
            DSlop1();
        }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            row = DSLOP.CurrentRow.Index;
            int a;
            a = DSLOP.RowCount;
            DSLOP.Rows[a - 1].Selected = true;
            DSlop1();
        }
        //-----------------------------------------------------------END OF CLASSES MANAGEMENT-------------------------------------------------------------


//BÁO CÁO
        private void sinhViênToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            if (txtName.Text != null)
            {
                Dictionary<string, string> templateCollection = new Dictionary<string, string>();
                templateCollection.Add("BaoCaoSV", @"\Data\BaoCaoSV.docx");
                var dataSource = new
                {
                    hoten = txtName.Text,
                    masv = txtMsv.Text,
                    ngaysinh = txtBornDate.Text,
                    gioitinh = cbxSex.Text,
                    quequan = txtQue.Text,
                    diemrl = txtDrl.Text,
                    malop = textBox2.Text,
                };

                foreach (KeyValuePair<string, string> template in templateCollection)
                {
                    DocumentCore dc = DocumentCore.Load(template.Value);

                    dc.MailMerge.Execute(dataSource);

                    string readyDocPath = String.Format("{0}-{1}.{2}", template.Key, dataSource.masv, ".docx");

                    // The file format is detected automatically from the file extension.
                    dc.Save(readyDocPath);

                    // Open the ready document for demonstration purposes.
                    System.Diagnostics.Process.Start(readyDocPath);
                }
            }
            else
            {
                MessageBox.Show("Chọn sinh viên để in báo cáo!");
            }
        }

        private void lớpToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            if (txtmalop.Text != null)
            {
                Dictionary<string, string> templateCollection = new Dictionary<string, string>();
                templateCollection.Add("BaoCaoLop", @"\Data\BaoCaoLop.docx");
                var dataSource = new
                {
                    malop = txtmalop.Text,
                    tenlop = txtTenlop.Text,
                    email = txtMaillop.Text,
                    hotenlt = txtTruonglop.Text,
                };

                foreach (KeyValuePair<string, string> template in templateCollection)
                {
                    DocumentCore dc = DocumentCore.Load(template.Value);

                    dc.MailMerge.Execute(dataSource);

                    string readyDocPath = String.Format("{0}-{1}.{2}", template.Key, dataSource.malop, ".docx");

                    // The file format is detected automatically from the file extension.
                    dc.Save(readyDocPath);

                    // Open the ready document for demonstration purposes.
                    System.Diagnostics.Process.Start(readyDocPath);
                }
            }
            else
            {
                MessageBox.Show("Chọn lớp để in báo cáo!");
            }

        }
        //-----------------------------------------------------------END OF REPORT-------------------------------------------------------------
        private void tool2_Click(object sender, EventArgs e)
        {
            groupBox1.BringToFront();
        }

        private void đổiMậtKhẩuToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Nhom2_ChangePass f = new Nhom2_ChangePass();
            f.Show();
        }

        private void thoátToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void sinhViênToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            groupBox2.BringToFront();
            groupBox3.Visible = false;
        }

        private void lớpToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            groupBox3.BringToFront();
            groupBox3.Visible = true;
        }

        private void wordToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Process.Start("WINWORD.EXE");
        }

        private void exelToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Process.Start("EXCEL.EXE");
        }

        private void paintToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Process.Start("mspaint.exe");
        }
        
        private void tácGiảToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Nhom2_Help n = new Nhom2_Help();
            n.panel1.BringToFront();
            n.panel2.BringToFront();
            n.panel3.BringToFront();
            n.panel4.BringToFront();
            n.Show();
        }

        private void bảnQuyềnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Nhom2_Help n = new Nhom2_Help();
            //n.panel2.Visible = false;
            //n.panel1.Visible = false;
            //n.panel3.Visible = false;
            //n.panel4.Visible = false;
            n.l5.Visible = true;
            n.l5.BringToFront();
            n.Show();
        }
    }
}
