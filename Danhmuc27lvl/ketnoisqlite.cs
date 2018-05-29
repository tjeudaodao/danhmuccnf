using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Spire.Xls;
using System.Drawing.Imaging;
using Microsoft.Office.Interop;
using System.Windows.Forms;

namespace Danhmuc27lvl
{
    class ketnoisqlite
    {
        #region khoitao
        public SQLiteConnection connec = null;
        public ketnoisqlite()
        {
            string chuoiketnoi = "Data Source=dbhangmoi.db;version=3;new=false";
            connec = new SQLiteConnection(chuoiketnoi);
        }
        
        private static ketnoisqlite _khoitao = null;
        public static ketnoisqlite khoitao()
        {
            if (_khoitao==null)
            {
                _khoitao = new ketnoisqlite();
            }
            return _khoitao;
        }
      
        public void Open()
        {
            if (connec.State != ConnectionState.Open)
            {
                connec.Open();
            }
        }
        public void Close()
        {
            if (connec.State != ConnectionState.Closed)
            {
                connec.Close();
            }
        }
        #endregion

        #region doc file excel
        string ngaychen = DateTime.Now.ToString("dd-MM-yyyy");
        public string Kiemtrafile(string tenfile)
        {
            string sql = string.Format("select name from filedanhmuc where name='{0}'", tenfile);
            string giatri = null;
            Open();
                SQLiteCommand cmd = new SQLiteCommand(sql, connec);
                SQLiteDataReader dtr = cmd.ExecuteReader();
                
                while (dtr.Read())
                {
                    giatri = dtr["name"].ToString();
                }
            Close();
            return giatri;
        }
        public void Chenvaobangfiledanhmuc(string tenfile)
        {
            string sqlchen = string.Format(@"INSERT INTO filedanhmuc VALUES('{0}','{1}')", tenfile, ngaychen);
            Open();
                SQLiteCommand cmd = new SQLiteCommand(sqlchen, connec);
                cmd.ExecuteNonQuery();
            Close();
        }
        public string Kiemtra(string cotcankiem,string tenbangkiem,string giatricantim)
        {
            string sql = string.Format("select {0} from {1} where name='{2}'", cotcankiem,tenbangkiem,giatricantim);
            string giatri = null;
            Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            SQLiteDataReader dtr = cmd.ExecuteReader();

            while (dtr.Read())
            {
                giatri = dtr[cotcankiem].ToString();
            }
            Close();
            return giatri;
        }
        public void Chenvaobanghangduocban(string maduocban,string ngayduocban)
        {
            string sqlchen = string.Format(@"INSERT INTO hangduocban VALUES('{0}','{1}')", maduocban, ngayduocban);
            Open();
            SQLiteCommand cmd = new SQLiteCommand(sqlchen, connec);
            cmd.ExecuteNonQuery();
            Close();
        }
        #endregion
    }
}
