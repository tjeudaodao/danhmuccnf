using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Tulpep.NotificationWindow;
using System.Globalization;
//using Timerhts = System.Windows.Forms.Timer;

namespace Danhmuc27lvl
{
    public partial class Formchinh : Form
    {
        Thread luongmail;
        Thread newmail;
        Thread xulyanh;
        Thread chenmahang;
        Thread chaynen;
        Thread tudongloadanh;
        Thread luongnenmail;

        string duongdanchuaanh = Application.StartupPath + @"\luuanh\";
        static string ngaychonbandau = null;
        int sodongchon = 0;
        static string thoigiancapnhat = null;
        static bool phathaykhongphat = true;
        List<string> thongtinmailmoi;


        private ManualResetEvent dieukhienthread = new ManualResetEvent(true);

        public Formchinh()
        {
            InitializeComponent();
            chaynen = new Thread(luongchaynen);
            chaynen.IsBackground = true;
            chaynen.Start();

            tudongloadanh = new Thread(hamtudongloadanh);
            tudongloadanh.IsBackground = true;
            tudongloadanh.Start();
        }
        
        void luongchaynen()
        {
            while (true)
            {
                Thread.Sleep(10000);
                newmail = new Thread(hamloadmailmoi);
                newmail.IsBackground = true;
                newmail.Start();

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

                luongnenmail = new Thread(luongchaynenbaomail);
                luongnenmail.IsBackground = true;
                luongnenmail.Start();

                Thread.Sleep(600000);
            }
        }

        void luongchaynenbaomail()
        {
            newmail.Join();
            Thread.Sleep(1000);
            lbtrangthai.Invoke(new MethodInvoker(delegate ()
            {
                if (thongtinmailmoi != null)
                {
                    while (true)
                    {
                        foreach (string tt in thongtinmailmoi)
                        {
                            lbtrangthai.Text = tt;
                            Thread.Sleep(1000);
                        }
                    }

                }
            }));
        }
        void chenma()
        {
            xulyanh.Join(); //ham chenma(thread chenmahang) se doi cho ham xulyanh chay xong moi chay
            var ham = hamtao.Khoitao();
            ham.xulymahang();
            thoigiancapnhat = DateTime.Now.ToString("HH:mm:ss");
            lbthongbaocapnhat.Invoke(new MethodInvoker(delegate ()
            {
                lbthongbaocapnhat.Text = "Đã cập nhật xong -"+" Lúc: "+ thoigiancapnhat;
                lbthongbaocapnhat.ForeColor = System.Drawing.Color.Navy;
            }));
            pbtrangthaicapnhat.Invoke(new MethodInvoker(delegate ()
            {
                pbtrangthaicapnhat.Image = Properties.Resources.ok;
            }));
            var con = ketnoisqlite.khoitao();
            ngaychonbandau = con.layngayganhat();
            datag1.Invoke(new MethodInvoker(delegate ()
            {
                datag1.DataSource = con.laythongtinngayganhat(ngaychonbandau);
            }));
           lbtongma.Invoke(new MethodInvoker(delegate(){
               lbtongma.Text = con.tongmatrongngay(ngaychonbandau);
           }));
            
        }
        void hamloadmailmoi()
        {
            var xulymail = layfileoutlook.Instance();
            thongtinmailmoi = new List<string>();
            thongtinmailmoi = xulymail.loadmailmoi();
           
            pbmail.Invoke(new MethodInvoker(delegate ()
            {
                if (thongtinmailmoi != null)
                {
                    pbmail.Image = Properties.Resources.newmail;
                }
                else
                {
                    pbmail.Image = Properties.Resources.mail;
                }
            }));
            lbbaomail.Invoke(new MethodInvoker(delegate ()
            {
                if (thongtinmailmoi != null)
                {
                    lbbaomail.Text = "Có Mail mới";
                }
                else
                {
                    lbbaomail.Text = "Thông báo:";
                }
            }));

            
        }
        void hamcapnhat()
        {
            newmail.Join();

            var xulyoutlook = layfileoutlook.Instance();
            var ham = hamtao.Khoitao();
            xulyoutlook.xuly();

            ham.luudanhmuchangmoi();
            lbthongbaocapnhat.Invoke(new MethodInvoker(delegate ()
            {
                lbthongbaocapnhat.Text = "Đang cập nhật";// cho load tung file save trong mail
                lbthongbaocapnhat.ForeColor = System.Drawing.Color.Crimson;
            }));
            pbtrangthaicapnhat.Invoke(new MethodInvoker(delegate()
            {
                pbtrangthaicapnhat.Image = Properties.Resources.loading;
            }));
        }
        void hamxulyanh()
        {
            luongmail.Join(); // ham xulyanh se doi cho thread luonggmail chay xong moi bat day chay
            var ham = hamtao.Khoitao();
            ham.xulyanh();

        }
        void hamtudongloadanh()
        {
            while (true)
            {
                Thread.Sleep(2000);
                string[] tonghopanh = Directory.GetFiles(Application.StartupPath + @"\luuanh\");
                for (int i = 0; i < tonghopanh.Length; i++)
                {
                    pbanhsanpham.Invoke(new MethodInvoker(delegate ()
                    {
                        pbanhsanpham.ImageLocation = tonghopanh[i];
                    }));
                    lbmahang.Invoke(new MethodInvoker(delegate ()
                    {
                        lbmahang.Text = Path.GetFileNameWithoutExtension(tonghopanh[i]);
                    }));

                    Thread.Sleep(2000);

                    dieukhienthread.WaitOne(Timeout.Infinite);
                }
            }
            
        }
        /// <summary>
        /// 
        /// cac ham phuc vu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Formchinh_Load(object sender, EventArgs e)
        {
            var con = ketnoisqlite.khoitao();
            ngaychonbandau = con.layngayganhat();
            datag1.DataSource = con.laythongtinngayganhat(ngaychonbandau);
            updatesoluongtrenbang();
        }
        void laythongtinvaolabel(string mahang)
        {
            var conlite = ketnoisqlite.khoitao();
            var ham = hamtao.Khoitao();
            List<laythongtin> laytt = new List<laythongtin>();
            laytt = conlite.loclaythongtin1ma(mahang);
            string kiemtra = conlite.Kiemtra("matong", "hangduocban", mahang);
            if (kiemtra != null)
            {
                foreach (laythongtin tt in laytt)
                {
                    lbmotasanpham.Text = tt.Motamaban + " - " + tt.Chudemaban + " - " + tt.Ghichu;
                    lbngayban.Text = tt.Ngayduocban;
                    lbduocbanhaychua.Text = "Được bán";
                    string trunghang = conlite.laythongtintrunghang(mahang);
                    if (trunghang == null)
                    {
                        lbdatrunghaychua.Text = "Chưa trưng bán";
                    }
                    else
                    {
                        lbdatrunghaychua.Text = trunghang;
                    }
                    loadanh(mahang);
                }
            }
            else if (kiemtra == null)
            {
                pbThemvaoduocban.Enabled = true;
                lbdatrunghaychua.Text = "Chưa trưng bán";
                lbduocbanhaychua.Text = "Chưa được bán";
            }

        }
        void loadanh(string tenanh)
        {
            if (File.Exists(duongdanchuaanh + tenanh + ".png"))
            {
                dungphatanh();
                pbanhsanpham.ImageLocation = duongdanchuaanh + tenanh + ".png";
                lbmahang.Text = tenanh;
                
            }
            else
            {
                pbanhsanpham.Image = Properties.Resources.bombs;
            }
        }
        
        void updatetrunghangthanhdatrung()
        {
            var con = ketnoisqlite.khoitao();
            if (datag1.SelectedRows.Count > 0)
            {
                string matong = null;
                foreach (DataGridViewRow row in datag1.SelectedRows)
                {
                    matong = row.Cells[0].Value.ToString();
                    con.updatedatrunghangthanhdatrung(matong);
                    sodongchon = datag1.SelectedRows.Count;
                    NotificationHts("Vừa cập nhật : " + sodongchon.ToString() + " mã hàng");
                }
                datag1.DataSource = con.laythongtinkhichonngay(ngaychonbandau);
            }

        }
        void updatetrunghangthanhchuatrung()
        {
            var con = ketnoisqlite.khoitao();
            if (datag1.SelectedRows.Count > 0)
            {
                string matong = null;
                foreach (DataGridViewRow row in datag1.SelectedRows)
                {
                    matong = row.Cells[0].Value.ToString();
                    con.updatetrunghangthanhchuatrung(matong);
                    sodongchon = datag1.SelectedRows.Count;
                    NotificationHts("Vừa cập nhật : " + sodongchon.ToString() + " mã hàng");
                }
                datag1.DataSource = con.laythongtinkhichonngay(ngaychonbandau);
            }

        }
        void NotificationHts(string noidung)
        {
            PopupNotifier pop = new PopupNotifier();
            pop.TitleText = "Thông báo";
            pop.ContentText = noidung;
            pop.Image = Properties.Resources.chancho;
            pop.IsRightToLeft = true;
            
            pop.Popup();
        }
        void NotificationHts(string noidung, string tieude)
        {
            PopupNotifier pop = new PopupNotifier();
            pop.TitleText = tieude;
            pop.ContentText = noidung;
            pop.Popup();
        }
        
        void updatesoluongtrenbang()
        {
            lbtongma.Text = datag1.Rows.Count.ToString();
        }
        #region Thao tac xu kien
        private void txtbarcode_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (!string.IsNullOrEmpty(txtbarcode.Text))
                    {
                        var consql = ketnoi.Instance();
                        string masp = consql.laymasp(txtbarcode.Text);
                        laythongtinvaolabel(masp);

                        lbmahang.Text = masp;
                        loadanh(masp);
                        txtbarcode.Clear();
                        txtbarcode.Focus();
                    }
                }
            }
            catch (Exception ex)
            {

                lbtrangthai.Text=ex.ToString();
            }
           
        }

        private void txtmatong_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var consqlite = ketnoisqlite.khoitao();
                datag1.DataSource = consqlite.loctheotenmatong(txtmatong.Text);
                string mau = @"\d{1}\w{2}\d{2}[SWAC]\d{3}";
                if (Regex.IsMatch(txtmatong.Text, mau))
                {
                    laythongtinvaolabel(txtmatong.Text);
                }
                updatesoluongtrenbang();
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
        }
        private void pbxoamatong_Click(object sender, EventArgs e)
        {
            txtmatong.Clear();
            txtmatong.Focus();
        }
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            try
            {
                var month = sender as MonthCalendar;
                DateTime ngaychon = month.SelectionStart;
                ngaychonbandau = month.SelectionStart.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                var con = ketnoisqlite.khoitao();
                datag1.DataSource = con.laythongtinkhichonngay(ngaychonbandau);
                updatesoluongtrenbang();
                dateTimePicker1.Value = ngaychon;
                dateTimePicker2.Value = ngaychon;
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
            
        }

        private void pbThemvaoduocban_Click(object sender, EventArgs e)
        {
            try
            {
                var con = ketnoisqlite.khoitao();
                if (con.Kiemtra("matong", "hangduocban", lbmahang.Text) == null && lbmahang.Text != "Mã hàng")
                {
                    DialogResult hoi = MessageBox.Show("Thêm mã " + lbmahang.Text + " vào danh sách được bán", "Thông báo", MessageBoxButtons.YesNo);
                    if (hoi == DialogResult.Yes)
                    {
                        // update ma hang vao danh sach duoc ban
                        con.themmamoivaodanhsachduocban(lbmahang.Text);
                        NotificationHts("Vừa thêm mã " + lbmahang.Text + " vào danh sách được bán");
                    }
                }
                else { MessageBox.Show("Mã đấy đã có trong danh sách được bán"); }
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
            
        }

        private void btndatrunghang_Click(object sender, EventArgs e)
        {
            try
            {
                updatetrunghangthanhdatrung();
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
            
        }

        private void btnchuatrunghang_Click(object sender, EventArgs e)
        {
            try
            {

                updatetrunghangthanhchuatrung();
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
        }
        private void datag1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow row = datag1.Rows[e.RowIndex];
                lbmahang.Text = row.Cells[0].Value.ToString();
                lbmotasanpham.Text = row.Cells[1].Value.ToString() + " - " + row.Cells[2].Value.ToString() + " - " + row.Cells[3].Value.ToString();
                lbngayban.Text = row.Cells[4].Value.ToString();
                lbdatrunghaychua.Text = row.Cells[5].Value.ToString();
                if (lbdatrunghaychua.Text == null || lbdatrunghaychua.Text=="")
                {
                    lbdatrunghaychua.Text = "Chưa trưng bán";
                }
                lbduocbanhaychua.Text = "Đã được bán";
                loadanh(lbmahang.Text);
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
            
        }
        private void btnXuatIn_Click(object sender, EventArgs e)
        {
            try
            {
                string ngaybatdau = dateTimePicker1.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                string ngayketthuc = dateTimePicker2.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                var ham = hamtao.Khoitao();
                ngaybatdau = ham.chuyendoingayvedangso(ngaybatdau);
                ngayketthuc = ham.chuyendoingayvedangso(ngayketthuc);
                var con = ketnoisqlite.khoitao();
                DataTable dt = new DataTable();
                dt = con.laythongtinkhoangngay(ngaybatdau, ngayketthuc);
                if (ham.Xuatfileexcel(dt, ngaybatdau, ngayketthuc))
                {
                    ham.taovainfileexcel(con.laythongtinIn(ngaybatdau, ngayketthuc));
                    PopupNotifier popexcel = new PopupNotifier();
                    popexcel.TitleText = "Thông báo";
                    popexcel.ContentText = "Vừa xuất file excel \nClick vào đây để mở file";
                    popexcel.IsRightToLeft = false;
                    popexcel.Image = Properties.Resources.excel;
                    popexcel.Click += Popexcel_Click;
                    popexcel.Popup();

                }
                
                
            }
            catch (Exception ex)
            {

                lbtrangthai.Text = ex.ToString();
            }
            
        }

        private void Popexcel_Click(object sender, EventArgs e)
        {
            var ham = hamtao.Khoitao();
            ham.mofileexcelvualuu();
        }

        private void pbphatanh_Click(object sender, EventArgs e)
        {
            if (phathaykhongphat)
            {
                pbphatanh.Image = Properties.Resources.play;
                phathaykhongphat = !phathaykhongphat;
                pbanhsanpham.Image = Properties.Resources.totoro1;
                lbmahang.Text = "Totoro";
                dieukhienthread.Reset();
            }
            else
            {
                pbphatanh.Image = Properties.Resources.pause;
                phathaykhongphat = !phathaykhongphat;
                dieukhienthread.Set();
            }
        }
        void dungphatanh()
        {
            if (phathaykhongphat)
            {
                pbphatanh.Image = Properties.Resources.play;
                phathaykhongphat = !phathaykhongphat;
                dieukhienthread.Reset();
            }
        }

        #endregion


    }
}
