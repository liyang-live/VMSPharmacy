using System;
using System.Data;
using System.Transactions;
using System.Linq;
using SubSonic;
using VNS.Libs;
using VNS.HIS.DAL;

using System.Text;

using SubSonic;
using NLog;
using System.Collections.Generic;
namespace VNS.HIS.BusRule.Classes
{
    public class KCB_KEDONTHUOC
    {
         private NLog.Logger log;
         public KCB_KEDONTHUOC()
        {
            log = LogManager.GetCurrentClassLogger();
        }
         
         public DataTable LayThuoctrongkho(int id_kho, int id_maloaithuoc, string KieuKho)
         {
             return SPs.ThuocTimkiemThuoctrongkho(id_kho, id_maloaithuoc, KieuKho).GetDataSet().Tables[0];
         }
        
       
        

    }
}
