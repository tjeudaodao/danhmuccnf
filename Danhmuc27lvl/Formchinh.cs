using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Danhmuc27lvl
{
    public partial class Formchinh : Form
    {
        Thread luongmail;
        Thread xulyanh;
        Thread chaynen;
        int i = 0;
        public Formchinh()
        {
            InitializeComponent();
            chaynen = new Thread(luongchaynen);
            chaynen.IsBackground = true;
            chaynen.Start();
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
                xulyanh.Join();
            }
        }
       
        void hamcapnhat()
        {
            var xulyoutlook = layfileoutlook.Instance();
            var ham = hamtao.Khoitao();
            xulyoutlook.xuly();
          
            ham.luudanhmuchangmoi();
            lbnhap.Invoke(new MethodInvoker(delegate ()
            {
                lbnhap.Text=i++.ToString();
            }));
            
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
           
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var con1 = ketnoisqlite.khoitao();
            con1.Open();
                string sql1 = "select * from filedanhmuc";

                SQLiteDataAdapter dta = new SQLiteDataAdapter(sql1, con1.connec);
                DataTable dt = new DataTable();
                dta.Fill(dt);
            
                dataGridView1.DataSource = dt;
            con1.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var con = ketnoisqlite.khoitao();
            string[] danhsachfile = Directory.GetFiles(Application.StartupPath + @"\filedanhmuc");
            Console.Write(danhsachfile.Length);
            label6.Text = danhsachfile.Length.ToString();
        }
    }
}
