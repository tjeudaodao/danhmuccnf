using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Danhmuc27lvl
{
    class laythongtin
    {
        private string ngayduocban;
        private string maduocban;
        private string motamaban;
        private string chudemaban;

        public string Ngayduocban
        {
            get { return ngayduocban; }
            set { ngayduocban = value; }
        }
        public string Maduocban
        {
            get { return maduocban; }
            set { maduocban = value; }
        }
        public string Motamaban
        {
            get { return motamaban; }
            set { motamaban = value; }
        }
        public string Chudemaban
        {
            get { return chudemaban; }
            set { chudemaban = value; }
        }

        public laythongtin(string ngayduocban,string maduocban,string motamaban,string chudemaban)
        {
            this.ngayduocban = ngayduocban;
            this.maduocban = maduocban;
            this.motamaban = motamaban;
            this.chudemaban = chudemaban;
        }
    }
}
