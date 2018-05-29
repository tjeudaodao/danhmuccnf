﻿using System;
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
        public void chenmotachudesanpham(string motasanpham,string chudesanpham,string matong)
        {
            string sql = string.Format("INSERT INTO mota(mota2,bst) VALUES({0},{1}) WHERE matong='{2}'", motasanpham, chudesanpham, matong);
            MySqlCommand cmd = new MySqlCommand(sql, connection);
            Open();
            cmd.ExecuteNonQuery();
            Close();
        }
        #endregion
    }
}
