using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Spire.Xls;
using excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop;
using System.Threading;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Danhmuc27lvl
{
    class hamtao
    {
        #region khoitao class
        public hamtao()
        {

        }
        private static hamtao _khoitao = null;
        public static hamtao Khoitao()
        {
            if (_khoitao==null)
            {
                _khoitao = new hamtao();
            }
            return _khoitao;
        }
        #endregion
        #region danhmuc
        string maungay = @"\d{2}/\d{2}/\d{4}";
        static List<laythongtin> luuthongtin = new List<laythongtin>();
        static List<string> danhsachfilechuaxuly = new List<string>();
        // ham chuyen doi dinh dang ngay tu string sang dang so co the + -
        public string chuyendoingayvedangso(string ngaydangDDMMYYYY)
        {
            DateTime dt = DateTime.ParseExact(ngaydangDDMMYYYY, "dd/MM/yyyy", null);
            return dt.ToString("yyyyMMdd");
        }
        
        public void luudanhmuchangmoi()
        {
           
            var con = ketnoisqlite.khoitao();
            string[] danhsachfile = Directory.GetFiles(Application.StartupPath + @"\filedanhmuc\");

            for (int i = 0; i < danhsachfile.Length; i++)
            {
                if (con.Kiemtrafile(danhsachfile[i]) == null)
                {

                    con.Chenvaobangfiledanhmuc(danhsachfile[i]);
                }

            }
        }
        public void xulyanh()
        {
            var con = ketnoisqlite.khoitao();
            danhsachfilechuaxuly = con.layfilechuaxuly();
            foreach (string file in danhsachfilechuaxuly)
            {
                copyanhvathongtin(file);
            }
            
        }
        public void xulymahang()
        {
            var con = ketnoisqlite.khoitao();
            var conmysql = ketnoi.Instance();
            foreach (laythongtin mahang in luuthongtin)
            {
                if (con.Kiemtra("matong","hangduocban",mahang.Maduocban) ==null)
                {
                    con.Chenvaobanghangduocban(mahang.Maduocban, mahang.Ngayduocban,mahang.Ghichu,mahang.Ngaydangso);
                    con.Chenhoacupdatebangmota(mahang.Maduocban, mahang.Motamaban, mahang.Chudemaban);
                    conmysql.chenmotachudesanpham(mahang.Motamaban, mahang.Chudemaban, mahang.Maduocban);
                }
            }
            luuthongtin.Clear();
            foreach (string file in danhsachfilechuaxuly)
            {
                con.thaydoitrangthaidakiemtra(file);
            }
            danhsachfilechuaxuly.Clear();
        }
        public void copyanhvathongtin(string filecanlay)
        {
            var excelApp = new excel.Application();
            var wb = excelApp.Workbooks.Open(filecanlay);
            var ws = (excel.Worksheet)wb.Worksheets[2];
            string duongdanluuanh = Application.StartupPath + @"\luuanh";
            int hangbatdau = 0;
            //lay ngay tu file excel roi chuyen doi sang dinh dang khac truoc khi insert vao database
            string ngayduocban = null;
            string ngaydangso = null;
            MatchCollection mat = Regex.Matches(ws.Cells[7, 1].value, maungay);
            foreach (Match m in mat)
            {
                ngayduocban = m.Value.ToString();
            }
            ngaydangso = chuyendoingayvedangso(ngayduocban);
            List<string> tenanh = new List<string>();
            foreach (var pic in ws.Pictures())
            {
                hangbatdau = pic.TopLeftCell.Row;
                luuthongtin.Add(new laythongtin(ngayduocban, ws.Cells[hangbatdau, 5].value, ws.Cells[hangbatdau, 6].value, ws.Cells[hangbatdau, 10].value, ws.Cells[hangbatdau, 11].value,ngaydangso));
                tenanh.Add(ws.Cells[hangbatdau, 5].value);
            }

            string[] manganh = tenanh.ToArray();
            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
            Marshal.FinalReleaseComObject(wb);

            Thread.Sleep(5);
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filecanlay);

            Worksheet sheet = workbook.Worksheets[1];
            if (!Directory.Exists(duongdanluuanh))
            {
                Directory.CreateDirectory(duongdanluuanh);
            }

            for (int i = 1; i < manganh.Length; i++)
            {
                Spire.Xls.ExcelPicture picture = sheet.Pictures[i];
                if (!File.Exists(duongdanluuanh + @"\" + manganh[i] + ".png"))
                {
                    picture.Picture.Save(duongdanluuanh + @"\" + manganh[i] + ".png", ImageFormat.Png);
                }
               

            }
            
            workbook.Dispose();
        }
        public void xuatfileexcel(DataTable dt,string ngaybatdau,string ngayketthuc)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel (.xlsx)|*.xlsx";
                saveDialog.FileName = "Thống kê hàng từ ngày - " + ngaybatdau + " đến ngày - "+ngayketthuc;
                if (saveDialog.ShowDialog() != DialogResult.Cancel)
                {
                    string exportFilePath = saveDialog.FileName;
                    var newFile = new FileInfo(exportFilePath);
                    using (var package = new ExcelPackage(newFile))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("hts");

                        worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                        worksheet.Column(1).AutoFit();
                        worksheet.Column(2).AutoFit();
                        worksheet.Column(3).AutoFit();
                        worksheet.Column(4).AutoFit();
                        worksheet.Column(5).AutoFit();

                        worksheet.Column(6).AutoFit();
                        package.Save();

                    }
                }
            }
        }
        #endregion
    }
}
