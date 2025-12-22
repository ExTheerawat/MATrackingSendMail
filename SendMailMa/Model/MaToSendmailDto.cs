using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMailMa.Model
{
   public class MaToSendmailDto
    {
        public string TRH_KEY { get; set; }
        public int TRH_SEQ { get; set; }
        public DateTime TRH_CREATE_DATE { get; set; }
        public string TRH_CREATE_BY { get; set; }
        public DateTime? TRH_UPDATE_DATE { get; set; }
        public string TRH_UPDATE_BY { get; set; }
        public string TRH_PO_NO { get; set; }
        public string TRH_CUS_NAME { get; set; }
        public string TRH_PROJECT_NAME { get; set; }
        public string TRH_PROJECT_CODE { get; set; }
        public string TRH_AE_NAME { get; set; }
        public string TRH_VEN_NAME { get; set; }
        public string TRH_STATUS { get; set; }
        public bool TRH_STATUS_VALUE { get; set; }
        public string TRH_REMARK { get; set; }
        public DateTime TRH_START_DATE { get; set; }
        public DateTime TRH_END_DATE { get; set; }
        public string TRD_CODE { get; set; }
        public string TRD_PRODUCT_ITEM { get; set; }
        public int? TRD_QTY { get; set; }
        public string TRD_WARRANTY { get; set; }
        public string TRD_WAR { get; set; }
        public string TRSN_CODE { get; set; }
        public string TRH_TYPE { get; set; }
    }
}
