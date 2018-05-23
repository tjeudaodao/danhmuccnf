using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using MySql.Data.MySqlClient;

namespace Danhmuc27lvl
{
    class ketnoi
    {
       
        private ketnoi()
        {
        }
        
        public string Password { get; set; }
        private MySqlConnection connection = null;
        public MySqlConnection Connection
        {
            get { return connection; }
        }

        private static ketnoi _instance = null;
        public static ketnoi Instance()
        {
            if (_instance == null)
                _instance = new ketnoi();
            return _instance;
        }

        public bool IsConnect()
        {
            if (Connection == null)
            {
                //string connstring = string.Format("Server=27.72.29.28;port=3306; database={0}; User Id=kho; password=1234", databaseName);
                string connstring = string.Format("Server=localhost;port=3306; database=cnf; User Id=hts; password=1211");
                connection = new MySqlConnection(connstring);
                connection.Open();
            }

            return true;
        }

        public void Close()
        {
            connection.Close();
        }
    }
}
