using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Transactions;
using VNS.HIS.DAL;
using VNS.Libs;
using SubSonic;
using NLog;
using VNS.Properties;
using System.Data;
namespace VNS.HIS.NGHIEPVU.THUOC
{

    public class Quaythuoc
    {
        private NLog.Logger log;
        public Quaythuoc()
        {
            log = NLog.LogManager.GetCurrentClassLogger();
        }
        
    }
}
