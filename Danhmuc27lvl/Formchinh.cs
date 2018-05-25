using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Threading;

namespace Danhmuc27lvl
{
    public partial class Formchinh : Form
    {
        System.Windows.Forms.Timer bodem1;
        Thread luongmail;
        int i = 0;
        public Formchinh()
        {
            InitializeComponent();
            bodem1 = new System.Windows.Forms.Timer();
            bodem1.Interval = 3000;
            bodem1.Tick += Bodem1_Tick;
            bodem1.Start();
        }

        private void Bodem1_Tick(object sender, EventArgs e)
        {
            luongmail = new Thread(hamcapnhat);
            luongmail.IsBackground = true;
            luongmail.Start();
        }
        void hamcapnhat()
        {
            var xulyoutlook = layfileoutlook.Instance();
            xulyoutlook.xuly();
            lbnhap.Invoke(new MethodInvoker(delegate ()
            {
                lbnhap.Text=i++.ToString();
            }));
        }
        private void Formchinh_Load(object sender, EventArgs e)
        {
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
           
        }
    }
}
