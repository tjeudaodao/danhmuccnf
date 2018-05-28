﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
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
        static List<string> danhsachfiledachen = new List<string>();
        public void luudanhmuchangmoi()
        {
           
            var con = ketnoisqlite.khoitao();
            string[] danhsachfile = Directory.GetFiles(Application.StartupPath + @"\filedanhmuc\");

            for (int i = 0; i < danhsachfile.Length; i++)
            {
                if (con.Kiemtrafile(danhsachfile[i]) == null)
                {

                    con.Chenvaobangfiledanhmuc(danhsachfile[i]);
                    danhsachfiledachen.Add(danhsachfile[i]);
                }

            }
        }
        public void xulyanh()
        {
            string[] mangfile = danhsachfiledachen.ToArray();
            for (int i = 0; i < mangfile.Length; i++)
            {
                copyanhvathongtin(mangfile[i]);
                Console.Write("Dng xu ly file" + mangfile[i]);
            }
            danhsachfiledachen.Clear();
        }
        public void copyanhvathongtin(string filecanlay)
        {
            var excelApp = new excel.Application();
            var wb = excelApp.Workbooks.Open(filecanlay);
            var ws = (excel.Worksheet)wb.Worksheets[2];
            string duongdanluuanh = Application.StartupPath + @"\luuanh";
            int hangbatdau = 0;

            List<string> tenanh = new List<string>();
            foreach (var pic in ws.Pictures())
            {
                hangbatdau = pic.TopLeftCell.Row;

                tenanh.Add(ws.Cells[hangbatdau, 5].value);
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
        #endregion
    }
}
