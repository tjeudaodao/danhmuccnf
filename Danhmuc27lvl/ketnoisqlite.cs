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

        #region doc file excel, update xu ly chay nen
        string ngaychen = DateTime.Now.ToString("dd/MM/yyyy");
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
            string sqlchen = string.Format(@"INSERT INTO filedanhmuc VALUES('{0}','{1}','Not')", tenfile, ngaychen);
            Open();
                SQLiteCommand cmd = new SQLiteCommand(sqlchen, connec);
                cmd.ExecuteNonQuery();
            Close();
        }
        public List<string> layfilechuaxuly()
        {
            List<string> ds = new List<string>();
            string sql = "select name from filedanhmuc where tinhtrang='Not'";
            Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            SQLiteDataReader dtr = cmd.ExecuteReader();
            while (dtr.Read())
            {
                ds.Add(dtr["name"].ToString());
            }
            Close();
            return ds;
        }
        public void thaydoitrangthaidakiemtra(string tenfile)
        {
            string sql = string.Format("UPDATE filedanhmuc SET tinhtrang='{0}' WHERE name='{1}'", "OK", tenfile);
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            Open();
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
        public void Chenvaobanghangduocban(string maduocban,string ngayduocban,string ghichu)
        {
            string sqlchen = string.Format(@"INSERT INTO hangduocban(matong,ngayban,ghichu) VALUES('{0}','{1}','{2}')", maduocban, ngayduocban,ghichu);
            Open();
            SQLiteCommand cmd = new SQLiteCommand(sqlchen, connec);
            cmd.ExecuteNonQuery();
            Close();
        }
        public void Chenhoacupdatebangmota(string matong,string mota,string chude)
        {
            string sql = null;
            if (Kiemtra("matong1","mota",matong)==null)
            {
                sql = string.Format("INSERT INTO mota(mota2,bst) VALUES('{0}','{1}') WHERE matong1='{3}'", mota, chude, matong);
            }
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            Open();
            cmd.ExecuteNonQuery();
            Close();
        }
        #endregion

        #region xu ly tren from
        //lay tat ca bang hang ban theo ten ma tong
        public DataTable loctheotenmatong(string matong)
        {
            string sql = string.Format("SELECT matong as 'Mã tổng',mota2 as 'Mô tả',bst as 'Chủ đề',ghichu as 'Ghi chú',ngayban as 'Ngày bán',trunghang as 'Trưng hàng' FROM hangduocban INNER JOIN mota ON hangduocban.matong=mota.matong1 where matong like '{0}'", matong);
            DataTable dt = new DataTable();
            Open();
            SQLiteDataAdapter dta = new SQLiteDataAdapter(sql, connec);
            dta.Fill(dt);
            Close();
            return dt;
        }
        // laythong tin gan vao list<>
        public List<laythongtin> loclaythongtin1ma(string matong)
        {
            string sql = string.Format("SELECT matong as 'Mã tổng',mota2 as 'Mô tả',bst as 'Chủ đề',ghichu as 'Ghi chú',ngayban as 'Ngày bán' FROM hangduocban INNER JOIN mota ON hangduocban.matong=mota.matong1 where matong = '{0}'", matong);
            List<laythongtin> laytt = new List<laythongtin>();
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            Open();
            SQLiteDataReader dtr = cmd.ExecuteReader();
            while (dtr.Read())
            {
                laytt.Add(new laythongtin(dtr[4].ToString(), dtr[0].ToString(), dtr[1].ToString(), dtr[2].ToString(), dtr[3].ToString()));
            }
            Close();
            return laytt;
        }
        // lay thong tin trung hang
        public string laythongtintrunghang(string matong)
        {
            string sql = string.Format("select trunghang from hangduocban where matong='{0}'", matong);
            string h = null;
            Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, connec);
            SQLiteDataReader dtr = cmd.ExecuteReader();
            while (dtr.Read())
            {
                h = dtr[0].ToString();
            }
            Close();
            return h;
        }
        // lay thong tin khi kich chon ngay
        public DataTable laythongtinkhichonngay(string ngaychon)
        {
            string sql = string.Format("SELECT matong as 'Mã tổng',mota2 as 'Mô tả',bst as 'Chủ đề',ghichu as 'Ghi chú',ngayban as 'Ngày bán',trunghang as 'Trưng hàng' FROM hangduocban INNER JOIN mota ON hangduocban.matong=mota.matong1 where ngayban = '{0}'", ngaychon);
            DataTable dt = new DataTable();
            Open();
            SQLiteDataAdapter dta = new SQLiteDataAdapter(sql, connec);
            dta.Fill(dt);
            Close();
            return dt;
        }
        // xuat bang khi chon khoang ngay cho viec xuat excel va in
        public DataTable laythongtinkhoangngay(string ngaybatday,string ngayketthuc)
        {
            string sql = string.Format("SELECT matong as 'Mã tổng',mota2 as 'Mô tả',bst as 'Chủ đề',ghichu as 'Ghi chú',ngayban as 'Ngày bán',trunghang as 'Trưng hàng' FROM hangduocban INNER JOIN mota ON hangduocban.matong=mota.matong1 where ngayban > '{0}' and ngayban < '{1}'", ngaybatday, ngayketthuc);
            DataTable dt = new DataTable();
            Open();
            SQLiteDataAdapter dta = new SQLiteDataAdapter(sql, connec);
            dta.Fill(dt);
            Close();
            return dt;
        }
        // xuatbang cho viec in chi lay 3 cot matong bst ngayban
        public DataTable laythongtinIn(string ngaybatdau,string ngayketthuc)
        {
            string sql = string.Format("SELECT matong as 'Mã tổng',bst as 'Chủ đề',ngayban as 'Ngày bán' FROM hangduocban INNER JOIN mota ON hangduocban.matong=mota.matong1 where ngayban > '{0}' and ngayban < '{1}'", ngaybatdau, ngayketthuc);
            DataTable dt = new DataTable();
            Open();
            SQLiteDataAdapter dta = new SQLiteDataAdapter(sql, connec);
            dta.Fill(dt);
            Close();
            return dt;
        }
        // update gia tri cot trung hang thanh " da trung hang"
        public void updatedatrunghang(string matong)
        {
            if (Kiemtra("trunghang","hangduocban",matong)==null)
            {
                string sql = string.Format("UPDATE hangduocban SET trunghang='{0}' WHERE matong='{1}'", "Đã Trưng Bán", matong);
                SQLiteCommand cmd = new SQLiteCommand(sql, connec);
                Open();
                cmd.ExecuteNonQuery();
                Close();
            }
            else if (Kiemtra("trunghang", "hangduocban", matong) == "Đã Trưng Bán")
            {
                string sql = string.Format("UPDATE hangduocban SET trunghang='{0}' WHERE matong='{1}'", null, matong);
                SQLiteCommand cmd = new SQLiteCommand(sql, connec);
                Open();
                cmd.ExecuteNonQuery();
                Close();
            }
        }
        #endregion
    }
}
