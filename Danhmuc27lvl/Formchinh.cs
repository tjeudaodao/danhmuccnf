using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Danhmuc27lvl
{
    public partial class Formchinh : Form
    {
        Thread luongmail;
        Thread xulyanh;
        Thread chenmahang;
        Thread chaynen;
        string duongdanchuaanh = Application.StartupPath + @"\luuanh\";
        int i = 0;
        public Formchinh()
        {
            InitializeComponent();
            chaynen = new Thread(luongchaynen);
            chaynen.IsBackground = true;
            // chaynen.Start();
        }
        void luongchaynen()
        {
            while (true)
            {
                Thread.Sleep(15000);
                luongmail = new Thread(hamcapnhat);
                luongmail.IsBackground = true;
                luongmail.Start();

                xulyanh = new Thread(hamxulyanh);
                xulyanh.IsBackground = true;
                xulyanh.Start();

                chenmahang = new Thread(chenma);
                chenmahang.IsBackground = true;
                chenmahang.Start();

                chenmahang.Join();
            }
        }
        void chenma()
        {
            xulyanh.Join();
            var ham = hamtao.Khoitao();
            ham.xulymahang();
        }
        void hamcapnhat()
        {
            var xulyoutlook = layfileoutlook.Instance();
            var ham = hamtao.Khoitao();
            xulyoutlook.xuly();

            ham.luudanhmuchangmoi();
            //lbnhap.Invoke(new MethodInvoker(delegate ()
            //{
            //    lbnhap.Text = i++.ToString();
            //}));

        }
        void hamxulyanh()
        {
            luongmail.Join();
            var ham = hamtao.Khoitao();
            ham.xulyanh();

        }

        private void Formchinh_Load(object sender, EventArgs e)
        {/*
            var con = ketnoi.Instance();
            
            if (con.IsConnect())
            {
                string sql = "select * from data";
                MySqlDataAdapter dta = new MySqlDataAdapter(sql, con.Connection);
                DataTable dt = new DataTable();
                dta.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();
            }
           */
            //string s = "17/06/2007";
            //DateTime d = DateTime.ParseExact(s, "dd/MM/yyyy", null);
            ////label7.Text = d.ToString("yyyy/MM/dd");
            //var ham = hamtao.Khoitao();
            //string s = "2017/12/11";
            //label7.Text = ham.chuyendoidinhdangngayveDDMMYYYYY(s);
        }
        void laythongtinvaolabel(string mahang)
        {
            var conlite = ketnoisqlite.khoitao();
            var ham = hamtao.Khoitao();
            List<laythongtin> laytt = new List<laythongtin>();
            laytt = conlite.loclaythongtin1ma(mahang);
            string kiemtra = conlite.Kiemtra("matong", "hangduocban", mahang);
            if (kiemtra !=null)
            {
                foreach (laythongtin tt in laytt)
                {
                    lbmotasanpham.Text = tt.Motamaban + " - " + tt.Chudemaban + " - " + tt.Ghichu;
                    lbngayban.Text = ham.chuyendoidinhdangngayveDDMMYYYYY(tt.Ngayduocban);
                    lbduocbanhaychua.Text = "Được bán";
                    string trunghang = conlite.laythongtintrunghang(mahang);
                    if (trunghang == null)
                    {
                        lbdatrunghaychua.Text = "Chưa trưng hàng";
                    }
                    else
                    {
                        lbdatrunghaychua.Text = trunghang;
                    }
                    loadanh(mahang);
                }
            }
            else if (kiemtra==null)
            {
                pbThemvaoduocban.Enabled = true;
                lbdatrunghaychua.Text = "Chưa trưng hàng";
                lbduocbanhaychua.Text = "Chưa được bán";
            }

        }
        void loadanh(string tenanh)
        {
            if (File.Exists(duongdanchuaanh+tenanh+".png"))
            {
                pbanhsanpham.ImageLocation = duongdanchuaanh + tenanh + ".png";
                lbmahang.Text = tenanh;
            }
        }

        #region Thao tac xu kien
        private void txtbarcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(txtbarcode.Text))
                {
                    var consql = ketnoi.Instance();
                    laythongtinvaolabel(consql.laymasp(txtbarcode.Text));
                    txtbarcode.Clear();
                    txtbarcode.Focus();
                }
            }
        }

        private void txtmatong_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var consqlite = ketnoisqlite.khoitao();
                datag1.DataSource = consqlite.loctheotenmatong(txtmatong.Text);
                string mau = @"\d{1}\w{2}\d{2}[SWAC]\d{3}";
                if (Regex.IsMatch(txtmatong.Text,mau))
                {
                    laythongtinvaolabel(txtmatong.Text);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            var month = sender as MonthCalendar;
            string ngaytim = month.SelectionStart.ToString("yyyy-MM-dd");
            var con = ketnoisqlite.khoitao();
            datag1.DataSource = con.laythongtinkhichonngay(ngaytim);
        }

        private void pbThemvaoduocban_Click(object sender, EventArgs e)
        {

        }

        private void pbUpdatematrung_Click(object sender, EventArgs e)
        {

        }

        private void pbChuatrung_Click(object sender, EventArgs e)
        {

        }
        private void btnXuatIn_Click(object sender, EventArgs e)
        {

        }
        #endregion

    }
}
