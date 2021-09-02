using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importarmeta
{
    class META2
    {

        public META2(int id, string icodusur, string inomevendedor, string Icodigo, string idescricao, double ivlvendaprev, string icliposiprev, string iqtvendaprev, string imixprev)
        {
            CODIGO = Icodigo;
            CODUSUR = icodusur;
            NOMEVENDEDOR = inomevendedor;
            DESC = idescricao;
            VLVENDAPREV = ivlvendaprev;
            CLIPOSPREV = icliposiprev;
            ID = id;
            QTVENDAPREV = iqtvendaprev;
            MIXPREV = imixprev;

        }
        public string CODUSUR { get; set; }
        public string CODIGO { get; set; }
        public string NOMEVENDEDOR { get; set; }
        public string DESC { get; set; }
        public double VLVENDAPREV { get; set; }
        public string CLIPOSPREV { get; set; }
        public int ID { get; set; }
        public string QTVENDAPREV { get; set; }
        public string MIXPREV { get; set; }

    }
}
