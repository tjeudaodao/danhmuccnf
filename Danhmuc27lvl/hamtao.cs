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
        string maungay = @"\d{2}/[\d{2},\d{1}]/\d{4}";
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
        public bool xulyanh()
        {
            bool kq = false;
            var con = ketnoisqlite.khoitao();
            danhsachfilechuaxuly = con.layfilechuaxuly();
            string mau = @"^KH tung hang";

            foreach (string file in danhsachfilechuaxuly)
            {
                if (file == null)
                {
                    kq = false;
                }
                else
                {
                    try
                    {
                        if (Path.GetExtension(file) == ".xlsx")
                        {
                            // Console.WriteLine(file);
                            ExcelPackage filechon = new ExcelPackage(new FileInfo(file));
                            //ExcelWorksheet ws = filechon.Workbook.Worksheets[1];
                            foreach (ExcelWorksheet ws in filechon.Workbook.Worksheets)
                            {
                                //Console.WriteLine(Convert.ToString(ws.Cells[7, 1].Text));
                                if (Regex.IsMatch(Convert.ToString(ws.Cells[7, 1].Text) ?? "", maungay))
                                {
                                    var sodong = ws.Dimension.End.Row;

                                    string ngayduocban = null;
                                    string ngaydangso = null;
                                    MatchCollection mat = Regex.Matches(Convert.ToString(ws.Cells[7, 1].Text) ?? "", maungay);
                                    //Console.WriteLine(ws.Cells[7,1].value);
                                    foreach (Match m in mat)
                                    {
                                        ngayduocban = m.Value.ToString();
                                    }
                                    if (Regex.IsMatch(ngayduocban,@"\d{2}/\d{1}/\d{4}"))
                                    {
                                        ngayduocban = ngayduocban.Substring(0, 3) + "0" + ngayduocban.Substring(3, 6);
                                    }
                                   // Console.WriteLine(ngayduocban);
                                    string mahang, mota, bst, ghichu;
                                    ngaydangso = chuyendoingayvedangso(ngayduocban);
                                    for (int i = 10; i < sodong; i++)
                                    {
                                        if (ws.Cells[i, 5].Value == null)
                                        {
                                            continue;
                                        }
                                        mahang = ws.Cells[i, 5].Value.ToString();
                                        mota = ws.Cells[i, 6].Value.ToString();
                                        bst = Convert.ToString(ws.Cells[i, 10].Value);
                                        ghichu = Convert.ToString(ws.Cells[i, 11].Value);
                                        luuthongtin.Add(new laythongtin(ngayduocban, mahang, mota, bst, ghichu, ngaydangso));
                                    }
                                }
                                //else Console.WriteLine("ko khop");
                            }
                            
                            filechon.Dispose();
                        }
                        else if (Path.GetExtension(file) == ".xls")
                        {
                            if (Regex.IsMatch(Path.GetFileName(file), mau))
                            {
                                copyanhKHtunghang(file);
                            }
                            else copyanhvathongtin(file);
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }
                    kq = true;
                }
            }
            return kq;
        }
        public void xulymahang()
        {
            var con = ketnoisqlite.khoitao();
            var conmysql = ketnoi.Instance();
            try
            {
                foreach (laythongtin mahang in luuthongtin)
                {
                   // Console.WriteLine(mahang.Maduocban);
                    if (conmysql.Kiemtra("matong", "hangduocban", mahang.Maduocban) == null)
                    {
                        // con.Chenvaobanghangduocban(mahang.Maduocban, mahang.Ngayduocban, mahang.Ghichu, mahang.Ngaydangso,mahang.Motamaban,mahang.Chudemaban);
                        try
                        {
                            conmysql.Chenvaobanghangduocban(mahang.Maduocban, mahang.Ngayduocban, mahang.Ghichu, mahang.Ngaydangso, mahang.Motamaban, mahang.Chudemaban);
                        }
                        catch (Exception)
                        {

                            continue;
                        }
                    }


                }
                luuthongtin.Clear();
            }
            catch (Exception)
            {

                throw;
            }
            
            foreach (string file in danhsachfilechuaxuly)
            {
                con.thaydoitrangthaidakiemtra(file);
            }
            danhsachfilechuaxuly.Clear();
        }
        public void copyanhvathongtin(string filecanlay)
        {
            var excelApp = new excel.Application();
           // Console.WriteLine(filecanlay);
            var wbs = excelApp.Workbooks;
            var wb = wbs.Open(filecanlay);
            //var ws = (excel.Worksheet)wb.Worksheets[2];
            string duongdanluuanh = Application.StartupPath + @"\luuanh";
            List<string> tenanh = new List<string>();
            foreach (excel.Worksheet ws in wb.Worksheets)
            {
                if (Regex.IsMatch(Convert.ToString(ws.Cells[7, 1].value2) ?? "", maungay))
                {
                    int hangbatdau = 0;
                    //lay ngay tu file excel roi chuyen doi sang dinh dang khac truoc khi insert vao database
                    string ngayduocban = null;
                    string ngaydangso = null;
                    //Console.WriteLine(ws.Cells[7, 1].value2);
                    //Console.WriteLine(Convert.ToString(ws.Cells[7, 1].value2));
                    MatchCollection mat = Regex.Matches(Convert.ToString(ws.Cells[7, 1].value2) ?? "", maungay);
                    //Console.WriteLine(ws.Cells[7,1].value);
                    foreach (Match m in mat)
                    {
                        ngayduocban = m.Value.ToString();
                    }
                    if (Regex.IsMatch(ngayduocban, @"\d{2}/\d{1}/\d{4}"))
                    {
                        ngayduocban = ngayduocban.Substring(0, 3) + "0" + ngayduocban.Substring(3, 6);
                    }
                    //Console.WriteLine(ngayduocban);
                    ngaydangso = chuyendoingayvedangso(ngayduocban);


                    string mahang, mota, bst, ghichu;
                    foreach (var pic in ws.Pictures())
                    {
                        hangbatdau = pic.TopLeftCell.Row;
                        if ((ws.Cells[hangbatdau, 5].value == null))
                        {
                            continue;
                        }
                        tenanh.Add(ws.Cells[hangbatdau, 5].value2.ToString());
                       //  Console.WriteLine(hangbatdau.ToString());

                    }
                    int lastRow = ws.Cells[ws.Rows.Count, 5].End(excel.XlDirection.xlUp).Row;
                    //Console.WriteLine(lastRow.ToString());
                    for (int i = 10; i < (lastRow + 5); i++)
                    {
                        if (ws.Cells[i, 5].value == null)
                        {
                            continue;
                        }
                        mahang = ws.Cells[i, 5].value2.ToString();
                        mota = ws.Cells[i, 6].value2.ToString();
                        bst = Convert.ToString(ws.Cells[i, 10].value2);
                        ghichu = Convert.ToString(ws.Cells[i, 11].value2);
                        //Console.WriteLine(mahang);
                        luuthongtin.Add(new laythongtin(ngayduocban, mahang, mota, bst, ghichu, ngaydangso));
                    }
                }
                
            }
            string[] manganh = tenanh.ToArray();
            wb.Close();
            wbs.Close();
            excelApp.Quit();
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(wbs);
            Marshal.FinalReleaseComObject(excelApp);

            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filecanlay);

               // Worksheet sheet = workbook.Worksheets[1];
                if (!Directory.Exists(duongdanluuanh))
                {
                    Directory.CreateDirectory(duongdanluuanh);
                }
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    for (int i = 1; i < manganh.Length; i++)
                    {
                        Spire.Xls.ExcelPicture picture = sheet.Pictures[i];
                        if (!File.Exists(duongdanluuanh + @"\" + manganh[i] + ".png"))
                        {
                            picture.Picture.Save(duongdanluuanh + @"\" + manganh[i] + ".png", ImageFormat.Png);
                        }
                        // Console.WriteLine(manganh[i]);

                    }
                    // Console.WriteLine(sheet.Pictures.Count);
                }

                workbook.Dispose();
            }
            catch (Exception)
            {

                return;
            }
            
        }
        public void copyanhKHtunghang(string filecanlay)
        {
            var excelApp = new excel.Application();
            var wbs = excelApp.Workbooks;
            var wb = wbs.Open(filecanlay);
           // var ws = (excel.Worksheet)wb.Worksheets[1];
            string duongdanluuanh = Application.StartupPath + @"\luuanh";
            int hangbatdau = 0;
            List<string> tenanh = new List<string>();
            foreach (excel.Worksheet ws in wb.Worksheets)
            {
                foreach (var pic in ws.Pictures())
                {
                    hangbatdau = pic.TopLeftCell.Row;

                    tenanh.Add(ws.Cells[hangbatdau, 2].value);
                }

            }
            string[] manganh = tenanh.ToArray();

            wb.Close();
            wbs.Close();
            excelApp.Quit();
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(wbs);
            Marshal.FinalReleaseComObject(excelApp);
            


            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filecanlay);

           // Worksheet sheet = workbook.Worksheets[0];
            if (!Directory.Exists(duongdanluuanh))
            {
                Directory.CreateDirectory(duongdanluuanh);
            }
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 1; i < manganh.Length; i++)
                {
                    Spire.Xls.ExcelPicture picture = sheet.Pictures[i];
                    if (!File.Exists(duongdanluuanh + @"\" + manganh[i] + ".png"))
                    {
                        picture.Picture.Save(duongdanluuanh + @"\" + manganh[i] + ".png", ImageFormat.Png);
                    }

                }
            }
            
            workbook.Dispose();
        }
        public bool Xuatfileexcel(DataTable dt, string ngaybatdau, string ngayketthuc,string tongma)
        {
            bool bl = true;
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                Random rd = new Random();
                int songaunhien = rd.Next(1, 100);
                saveDialog.Filter = "Excel (.xlsx)|*.xlsx";
                saveDialog.FileName = "Thống kê hàng từ ngày - " + ngaybatdau + " đến ngày - " + ngayketthuc+" -vs"+songaunhien.ToString();
                if (saveDialog.ShowDialog() != DialogResult.Cancel)
                {
                    string exportFilePath = saveDialog.FileName;
                    duongdanluufileexcel = exportFilePath;
                    var newFile = new FileInfo(exportFilePath);
                    using (var package = new ExcelPackage(newFile))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("hts");
                        worksheet.Cells["A1"].Value = "Tổng số mã: " + tongma;
                        worksheet.Cells["A3"].LoadFromDataTable(dt, true);
                        worksheet.Column(1).AutoFit();
                        worksheet.Column(2).AutoFit();
                        worksheet.Column(3).AutoFit();
                        worksheet.Column(4).AutoFit();
                        worksheet.Column(5).AutoFit();

                        worksheet.Column(6).AutoFit();
                        package.Save();
                        package.Dispose();
                    }
                    bl = true;
                }
                else
                {
                    bl = false;
                }
            }
            return bl;
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
        public void taovainfileexcel(DataTable dt,string tongma)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet worksheet = ExcelPkg.Workbook.Worksheets.Add("hts");

            worksheet.Cells["A1"].Value = "Tổng mã:";
            worksheet.Cells["C1"].Value = tongma;
            worksheet.Cells["A3"].LoadFromDataTable(dt, true);

            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 13;
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
            sh.PrintOutEx();
            app.Quit();
            Marshal.FinalReleaseComObject(app);
            Marshal.FinalReleaseComObject(book);
        }
        #endregion
    }
}
