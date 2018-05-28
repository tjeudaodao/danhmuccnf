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
        System.Windows.Forms.Timer bodem1;
        System.Windows.Forms.Timer bodem2;
        Thread luongmail;
        Thread xulyanh;

        int i = 0;
        public Formchinh()
        {
            InitializeComponent();
            bodem1 = new System.Windows.Forms.Timer();
            bodem1.Interval = 10000;
            bodem1.Tick += Bodem1_Tick;
            bodem1.Start();

            bodem2 = new System.Windows.Forms.Timer();
            bodem2.Interval = 15000;
            bodem2.Tick += Bodem2_Tick;
            bodem2.Start();
        }

        private void Bodem2_Tick(object sender, EventArgs e)
        {
            
            xulyanh = new Thread(hamxulyanh);
            xulyanh.IsBackground = true;
            xulyanh.Start();
            bodem1.Start();
        }

        private void Bodem1_Tick(object sender, EventArgs e)
        {
            
            luongmail = new Thread(hamcapnhat);
            luongmail.IsBackground = true;
            luongmail.Start();
            bodem1.Stop();
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
