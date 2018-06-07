using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;

namespace Danhmuc27lvl
{
    class ketnoi
    {
        #region khoitao
        private ketnoi()
        {
            string connstring = string.Format("Server=27.72.29.28;port=3306; database=cnf; User Id=kho; password=1234");
            // string connstring = string.Format("Server=localhost;port=3306; database=cnf; User Id=hts; password=1211");
            connection = new MySqlConnection(connstring);
        }
        
        private MySqlConnection connection = null;
       
        private static ketnoi _instance = null;
        public static ketnoi Instance()
        {
            if (_instance == null)
                _instance = new ketnoi();
            return _instance;
        }
        public void Open()
        {
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }
        }
      
        public void Close()
        {
            if (connection.State != ConnectionState.Closed)
            {
                connection.Close();
            }
        }
        #endregion
        #region thao tac tren csdl mysql
        //kiem tra xem ma hang day co trong bang mota chua
        public string Kiemtra(string mahang)
        {
            string sql = @"SELECT matong1 FROM mota WHERE matong1='" + mahang + "'";
            MySqlCommand cmd = new MySqlCommand(sql, connection);
            string hh = null;
            Open();
            MySqlDataReader dtr = cmd.ExecuteReader();
            while (dtr.Read())
            {
                hh = dtr["matong1"].ToString();
            }
            Close();
            return hh;
        }

        // chen mota sp
        //public void chenmotachudesanpham(string motasanpham,string chudesanpham,string matong)
        //{
        //    if (Kiemtra(matong)==null)
        //    {
        //        string sql = @"INSERT INTO mota(mota2,bst) VALUES('"+motasanpham+"','"+chudesanpham+"')";
        //        MySqlCommand cmd = new MySqlCommand(sql, connection);
        //        Open();
        //        cmd.ExecuteNonQuery();
        //        Close();
        //    }
            
        //}
        // lay masp tu barcode
        public string laymasp(string barcode)
        {
            string sql = string.Format("SELECT masp FROM data WHERE barcode='{0}'", barcode);
            string h = null;
            MySqlCommand cmd = new MySqlCommand(sql, connection);
            Open();
            MySqlDataReader dtr = cmd.ExecuteReader();
            while (dtr.Read())
            {
                h = dtr["masp"].ToString();
            }
            Close();
            int vitri = h.IndexOf("-");
            h = h.Substring(0, vitri);
            return h;
        }
        #endregion
    }
}
