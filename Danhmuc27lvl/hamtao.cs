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
        string maungay = @"\d{2}/(\d{2}|\d{1})/\d{4}";
        static List<laythongtin> luuthongtin = new List<laythongtin>();
        static List<string> danhsachfilechuaxuly = new List<string>();
        static string duongdanluufileexcel = null;

        string duongdanluuanh = Application.StartupPath + @"\luuanh";

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
            string mau = @"(^KH tung hang)|(^Ke hoach tung hang)";

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
                        if (Regex.IsMatch(Path.GetFileName(file), mau))
                        {
                            copyanhKHtunghang(file);
                        }
                        else copyanhvathongtin(file);

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
                    if (conmysql.Kiemtra("matong", "hangduocban", mahang.Maduocban) == null)
                    {
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
            catch (Exception ex)
            {
                ghiloi.WriteLogError(ex);
                return;
            }
            
            foreach (string file in danhsachfilechuaxuly)
            {
                con.thaydoitrangthaidakiemtra(file);
            }
            danhsachfilechuaxuly.Clear();
        }
        public void copyanhvathongtin(string filecanlay)
        {

            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filecanlay);
                if (!Directory.Exists(duongdanluuanh))
                {
                    Directory.CreateDirectory(duongdanluuanh);
                }
                foreach (Worksheet sheet in workbook.Worksheets)
                { 
                    if (Regex.IsMatch(sheet.Range[7, 1].Value2.ToString(), maungay))
                    {
                        string ngayduocban = null;
                        string ngaydangso = null;
                        Match layngay = Regex.Match(sheet.Range[7, 1].Value2.ToString(), maungay);
                        ngayduocban = layngay.Value;
                        if (Regex.IsMatch(ngayduocban, @"\d{2}/\d{1}/\d{4}"))
                        {
                            ngayduocban = ngayduocban.Substring(0, 3) + "0" + ngayduocban.Substring(3, 6);
                        }
                        ngaydangso = chuyendoingayvedangso(ngayduocban);

                        string mahang, mota, bst, ghichu;
                        int dongcuoi = sheet.LastRow;
                        for (int i = sheet.LastRow; i >= 0; i--)
                        {
                            CellRange cr = sheet.Rows[i - 1].Columns[1];
                            if (!cr.IsBlank)
                            {
                                dongcuoi = i;
                                break;
                            }
                        }
                        for (int i = 10; i < dongcuoi; i++)
                        {
                            if (sheet.Range[i, 5].Value == null || sheet.Range[i,5].Value == "")
                            {
                                continue;
                            }
                            mahang = sheet.Range[i, 5].Value2.ToString();
                            mota = sheet.Range[i, 6].Value2.ToString();
                            bst = sheet.Range[i, 10].Value2.ToString();
                            ghichu = sheet.Range[i, 11].Value2.ToString();
                            luuthongtin.Add(new laythongtin(ngayduocban, mahang, mota, bst, ghichu, ngaydangso));
                        }
                        string tenanhH = "";
                        for (int i = 1; i < sheet.Pictures.Count; i++)
                        {
                            Spire.Xls.ExcelPicture anh = sheet.Pictures[i];
                            tenanhH = sheet.Range[anh.TopRow, 5].Value2.ToString();
                            if (!File.Exists(duongdanluuanh + @"\" + tenanhH + ".png"))
                            {
                                anh.Picture.Save(duongdanluuanh + @"\" + tenanhH + ".png", ImageFormat.Png);
                            }
                        }
                    }
                    
                }

                workbook.Dispose();
            }
            catch (Exception ex)
            {
                ghiloi.WriteLogError(ex);
                return;
            }
            
        }
        public void copyanhKHtunghang(string filecanlay)
        {
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filecanlay);
                if (!Directory.Exists(duongdanluuanh))
                {
                    Directory.CreateDirectory(duongdanluuanh);
                }
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Pictures.Count > 0)
                    {
                        string tenanhH = "";
                        for (int i = 1; i < sheet.Pictures.Count; i++)
                        {
                            Spire.Xls.ExcelPicture anh = sheet.Pictures[i];
                            tenanhH = sheet.Range[anh.TopRow, 5].Value2.ToString();
                            if (!File.Exists(duongdanluuanh + @"\" + tenanhH + ".png"))
                            {
                                anh.Picture.Save(duongdanluuanh + @"\" + tenanhH + ".png", ImageFormat.Png);
                            }
                        }
                    }
                }

                workbook.Dispose();
            }
            catch (Exception e)
            {
                ghiloi.WriteLogError(e);
                return;
            }
            
        }

        // method cho in va xuat excel
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
        public void taovainfileexcel(DataTable dt, string tongma, string ngaybatdau, string ngayketthuc)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet worksheet = ExcelPkg.Workbook.Worksheets.Add("hts");

            worksheet.Cells["A1:C1"].Merge = true;
            worksheet.Cells["A2:C2"].Merge = true;
            worksheet.Cells["A3:C3"].Merge = true;
            worksheet.Cells["A1"].Value = "Danh mục VM _ Mã chưa trưng";
            worksheet.Cells["A2"].Value = "Từ ngày: " + ngaybatdau + " - " + ngayketthuc;

            worksheet.Cells["A3"].Value = "Tổng mã chưa trưng: " + tongma;
            worksheet.Cells["A5"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light1);

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
