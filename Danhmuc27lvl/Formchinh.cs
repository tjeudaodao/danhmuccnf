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

namespace Danhmuc27lvl
{
    public partial class Formchinh : Form
    {
        
        public Formchinh()
        {
            InitializeComponent();
           
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
           // var xulyoutlook = layfileoutlook.Instance();
            //xulyoutlook.xuly();
        }
    }
}
