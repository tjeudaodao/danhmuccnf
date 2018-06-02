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
using System.Drawing;

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
            if (_khoitao == null)
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
        static string duongdanluufileexcel = null;

        // ham chuyen doi dinh dang ngay tu string sang dang so co the + -
        public string chuyendoingayvedangso(string ngaydangDDMMYYYY)
        {
            try
            {
                DateTime dt = DateTime.ParseExact(ngaydangDDMMYYYY, "dd/MM/yyyy", null);
                return dt.ToString("yyyyMMdd");
            }
            catch (Exception)
            {

                return "Loi";
            }
            
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
                //Console.WriteLine(file);
                copyanhvathongtin(file);
            }

        }
        public void xulymahang()
        {
            var con = ketnoisqlite.khoitao();
            var conmysql = ketnoi.Instance();
            foreach (laythongtin mahang in luuthongtin)
            {
                if (con.Kiemtra("matong", "hangduocban", mahang.Maduocban) == null)
                {
                    con.Chenvaobanghangduocban(mahang.Maduocban, mahang.Ngayduocban, mahang.Ghichu, mahang.Ngaydangso,mahang.Motamaban,mahang.Chudemaban);
                    try
                    {
                        conmysql.chenmotachudesanpham(mahang.Motamaban, mahang.Chudemaban, mahang.Maduocban);
                    }
                    catch (Exception)
                    {

                        continue;
                    }
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
            //Console.WriteLine(ws.Cells[7,1].value);
            foreach (Match m in mat)
            {
                ngayduocban = m.Value.ToString();
            }
            Console.WriteLine(ngayduocban);
            ngaydangso = chuyendoingayvedangso(ngayduocban);
            
            List<string> tenanh = new List<string>();
            string mahang, mota, bst, ghichu;
            foreach (var pic in ws.Pictures())
            {
                hangbatdau = pic.TopLeftCell.Row;
                if (hangbatdau > 1)
                {
                    mahang=ws.Cells[hangbatdau, 5].value2.ToString();
                    mota = ws.Cells[hangbatdau, 6].value2.ToString();
                    bst=ws.Cells[hangbatdau, 10].value2.ToString();
                    ghichu=Convert.ToString(ws.Cells[hangbatdau, 11].value2);

                    luuthongtin.Add(new laythongtin(ngayduocban, mahang, mota,bst , ghichu, ngaydangso));
                    tenanh.Add(ws.Cells[hangbatdau, 5].value);
                }

            }

            string[] manganh = tenanh.ToArray();
            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
            Marshal.FinalReleaseComObject(wb);

           
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
        public void xuatfileexcel(DataTable dt, string ngaybatdau, string ngayketthuc)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel (.xlsx)|*.xlsx";
                saveDialog.FileName = "Thống kê hàng từ ngày - " + ngaybatdau + " đến ngày - " + ngayketthuc;
                if (saveDialog.ShowDialog() != DialogResult.Cancel)
                {
                    string exportFilePath = saveDialog.FileName;
                    duongdanluufileexcel = exportFilePath;
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
                        package.Dispose();
                    }
                }
            }
        }
        public void mofileexcelvualuu()
        {
            if (duongdanluufileexcel!=null)
            {
                var app = new excel.Application();

                excel.Workbooks book = app.Workbooks;
                excel.Workbook sh = book.Open(duongdanluufileexcel);
                app.Visible = true;
                //sh.PrintOutEx();
            }
            
        }
        public void taovainfileexcel(DataTable dt)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet worksheet = ExcelPkg.Workbook.Worksheets.Add("hts");
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);

            worksheet.Column(1).Width = 11;
            worksheet.Column(2).Width = 10;
            worksheet.Column(3).Width = 10;


            //worksheet.Cells[worksheet.Dimension.End.Row + 1, 1].Value = "Tổng sản phẩm:";
            //worksheet.Cells[worksheet.Dimension.End.Row, 2].Value = tongsp;

            var allCells = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 10));

            worksheet.PrinterSettings.LeftMargin = 0.2M / 2.54M;
            worksheet.PrinterSettings.RightMargin = 0.2M / 2.54M;
            worksheet.PrinterSettings.TopMargin = 0.2M / 2.54M;
            worksheet.Protection.IsProtected = false;
            worksheet.Protection.AllowSelectLockedCells = false;
            if (File.Exists("hts.xlsx"))
            {
                File.Delete("hts.xlsx");

            }
            ExcelPkg.SaveAs(new FileInfo("hts.xlsx"));
            ExcelPkg.Dispose();

            var app = new excel.Application();

            excel.Workbooks book = app.Workbooks;
            excel.Workbook sh = book.Open(Path.GetFullPath("hts.xlsx"));
            //app.Visible = true;
            //sh.PrintOutEx();
            app.Quit();
            Marshal.FinalReleaseComObject(app);
            Marshal.FinalReleaseComObject(book);
        }
        #endregion
    }
}
