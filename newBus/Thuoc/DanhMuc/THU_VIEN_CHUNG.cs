using System;
using System.Data;

using System.Linq;
using SubSonic;
using VNS.Libs;
using VNS.HIS.DAL;

using System.Text;

using SubSonic;
using NLog;
using System.Net.NetworkInformation;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Transactions;
using System.Windows.Forms;
using VNS.Properties;

namespace VNS.Libs
{
    public class THU_VIEN_CHUNG
    {
       
        public static int Songay(DateTime from,DateTime to)
        {
            int Tinh1Ngayneunhohon24h = Utility.Int32Dbnull(Laygiatrithamsohethong("NOITRU_NHOHON24H_TINH1NGAY", "1", false), 1);
            int Sogiotinh = Utility.Int32Dbnull(Laygiatrithamsohethong("NOITRU_SOGIO_LAMTRONNGAY", "1", false), 1);
            double totalhour =Math.Ceiling( (to - from).TotalHours);
            if(totalhour<24 )
            {
                if (Tinh1Ngayneunhohon24h==1) return 1;
                else if (totalhour >= Sogiotinh) return 1;
                else return 0;
            }
            int songay =(int) totalhour / 24;
            int sogio = (int)totalhour % 24;
            int songaythem=0;
            if (sogio >= Sogiotinh) songaythem = 1;

            return songay + songaythem;
        }
       
        public static string Laygiatrithamsohethong(string ParamName, bool fromDB)
        {
            try
            {
                string reval = null;
                if (fromDB)
                {
                    SqlQuery sqlQuery =
                        new Select().From(SysSystemParameter.Schema).Where(SysSystemParameter.Columns.SName).IsEqualTo(
                            ParamName);
                    SysSystemParameter objSystemParameter = sqlQuery.ExecuteSingle<SysSystemParameter>();
                    if (objSystemParameter != null) reval = objSystemParameter.SValue;
                }
                else
                {
                    DataRow[] arrDR = globalVariables.gv_dtSysparams.Select(SysSystemParameter.SNameColumn.ColumnName + " ='" + ParamName + "'");
                    if (arrDR.Length > 0) reval = Utility.sDbnull(arrDR[0][SysSystemParameter.SValueColumn.ColumnName]);
                }
                return reval;
            }
            catch
            {
                return null;
            }
        }
        public static SysSystemParameter Laythamsohethong(string ParamName)
        {
            try
            {

                SqlQuery sqlQuery =
                    new Select().From(SysSystemParameter.Schema).Where(SysSystemParameter.Columns.SName).IsEqualTo(
                        ParamName);
                SysSystemParameter objSystemParameter = sqlQuery.ExecuteSingle<SysSystemParameter>();

                return objSystemParameter;
            }
            catch
            {
                return null;
            }
        }
      
        public static QheDoituongThuoc LayQheDoituongThuoc(string madoituong, int idthuoc, string makhoathuchien, bool THUOC_GIATHEO_KHOAKCB)
        {
            SqlQuery sqlQuery = new Select().From(QheDoituongThuoc.Schema)
                                       .Where(QheDoituongThuoc.Columns.IdThuoc).IsEqualTo(
                                           idthuoc)
                                       .And(QheDoituongThuoc.Columns.MaDoituongKcb).IsEqualTo(madoituong);
            if (THUOC_GIATHEO_KHOAKCB)
                sqlQuery.And(QheDoituongThuoc.Columns.MaKhoaThuchien).IsEqualTo(makhoathuchien);
            QheDoituongThuoc objQheDoituongThuoc = sqlQuery.ExecuteSingle<QheDoituongThuoc>();
            return objQheDoituongThuoc;
        }
        public static void CapnhatgiatriTieudebaocao(string Matieude, string _value)
        {
            try
            {
                if (Utility.DoTrim(Matieude) == "") return;
                DataRow[] arrDR = globalVariables.gv_dtSysTieude.Select(SysTieude.MaTieudeColumn.ColumnName + " ='" + Matieude + "'");
                SysReport _Item = new Select().From(SysReport.Schema).Where(SysReport.Columns.MaBaocao).IsEqualTo(
                            Matieude).ExecuteSingle<SysReport>();
                if (_Item!=null)
                {
                    //arrDR[0][SysTieude.NoiDungColumn.ColumnName] = _value;
                    //globalVariables.gv_dtSysTieude.AcceptChanges();
                    new Update(SysReport.Schema).Set(SysReport.TieuDeColumn).EqualTo(_value).Where(SysReport.MaBaocaoColumn).IsEqualTo(Matieude).Execute();
                }
                else
                {
                    SysReport newItem = new SysReport();
                    newItem.MaBaocao = Matieude;
                    newItem.TieuDe = _value;

                    newItem.Save();
                    //DataRow newrow = globalVariables.gv_dtSysTieude.NewRow();
                    //newrow[SysTieude.MaTieudeColumn.ColumnName] = Matieude;
                    //newrow[SysTieude.NoiDungColumn.ColumnName] = _value;

                    //globalVariables.gv_dtSysTieude.Rows.Add(newrow);
                    //globalVariables.gv_dtSysTieude.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg("Lỗi khi cập nhật giá trị tiêu đề báo cáo:\n" + ex.Message);
            }
        }
        public static string LaygiatriTieudebaocao(string Matieude, string defaultval, bool fromDB)
        {
            try
            {
                string reval = defaultval;
                if (fromDB)
                {
                    SqlQuery sqlQuery =
                        new Select().From(SysReport.Schema).Where(SysReport.Columns.MaBaocao).IsEqualTo(
                            Matieude);
                    SysReport objSystemParameter = sqlQuery.ExecuteSingle<SysReport>();
                    if (objSystemParameter != null) reval = objSystemParameter.TieuDe;
                }
                else
                {
                    DataRow[] arrDR = globalVariables.gv_dtSysTieude.Select(SysTieude.MaTieudeColumn.ColumnName + " ='" + Matieude + "'");
                    if (arrDR.Length > 0) reval = Utility.sDbnull(arrDR[0][SysTieude.NoiDungColumn.ColumnName]);
                }
                return reval;
            }
            catch
            {
                return defaultval;
            }
        }
        public static string LaygiatriTieudeBaocao(string Matieude, bool fromDB)
        {
            try
            {
                string reval = "";
                if (fromDB)
                {
                    SqlQuery sqlQuery =
                        new Select().From(SysReport.Schema).Where(SysReport.Columns.MaBaocao).IsEqualTo(
                            Matieude);
                    SysReport objSystemParameter = sqlQuery.ExecuteSingle<SysReport>();
                    if (objSystemParameter != null) reval = objSystemParameter.TieuDe;
                }
                else
                {
                    DataRow[] arrDR = globalVariables.gv_dtSysTieude.Select(SysTieude.MaTieudeColumn.ColumnName + " ='" + Matieude + "'");
                    if (arrDR.Length > 0) reval = Utility.sDbnull(arrDR[0][SysTieude.NoiDungColumn.ColumnName]);
                }
                return reval;
            }
            catch
            {
                return "";
            }
        }
        public static void Capnhatgiatrithamsohethong(string ParamName, string _value)
        {
            try
            {
                if (Utility.DoTrim(ParamName) == "") return;
                DataRow[] arrDR = globalVariables.gv_dtSysparams.Select(SysSystemParameter.SNameColumn.ColumnName + " ='" + ParamName + "'");
                if (arrDR.Length > 0)
                {
                    arrDR[0][SysSystemParameter.SValueColumn.ColumnName] = _value;
                    globalVariables.gv_dtSysparams.AcceptChanges();
                    new Update(SysSystemParameter.Schema).Set(SysSystemParameter.SValueColumn).EqualTo(_value).Where(SysSystemParameter.SNameColumn).IsEqualTo(ParamName).Execute();
                }
                else
                {
                    SysSystemParameter newItem = new SysSystemParameter();
                    newItem.FpSBranchID = globalVariables.Branch_ID;
                    newItem.SName = ParamName;
                    newItem.SValue = _value;
                    newItem.IMonth = 0;
                    newItem.IYear = 0;
                    newItem.IStatus = 1;
                    newItem.IsNew = true;
                    newItem.Save();
                    DataRow newrow = globalVariables.gv_dtSysparams.NewRow();
                    newrow[SysSystemParameter.FpSBranchIDColumn.ColumnName] = globalVariables.Branch_ID;
                    newrow[SysSystemParameter.SNameColumn.ColumnName] = ParamName;
                    newrow[SysSystemParameter.SValueColumn.ColumnName] = _value;
                    newrow[SysSystemParameter.IYearColumn.ColumnName] = 0;
                    newrow[SysSystemParameter.IMonthColumn.ColumnName] = 0;
                    newrow[SysSystemParameter.IStatusColumn.ColumnName] = 1;
                    globalVariables.gv_dtSysparams.Rows.Add(newrow);
                    globalVariables.gv_dtSysparams.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg("Lỗi khi cập nhật giá trị tham số hệ thống:\n" + ex.Message);
            }
        }
        public static void Capnhatgiatrithamsohethong(string ParamName, string _value,string sdesc)
        {
            try
            {
                if (Utility.DoTrim(ParamName) == "") return;
                DataRow[] arrDR = globalVariables.gv_dtSysparams.Select(SysSystemParameter.SNameColumn.ColumnName + " ='" + ParamName + "'");
                if (arrDR.Length > 0)
                {
                    arrDR[0][SysSystemParameter.SValueColumn.ColumnName] = _value;
                    arrDR[0][SysSystemParameter.SDescColumn.ColumnName] = sdesc;
                    globalVariables.gv_dtSysparams.AcceptChanges();
                    new Update(SysSystemParameter.Schema).Set(SysSystemParameter.SValueColumn).EqualTo(_value).Set(SysSystemParameter.SDescColumn).EqualTo(sdesc).Where(SysSystemParameter.SNameColumn).IsEqualTo(ParamName).Execute();
                }
                else
                {
                    SysSystemParameter newItem = new SysSystemParameter();
                    newItem.FpSBranchID = globalVariables.Branch_ID;
                    newItem.SName = ParamName;
                    newItem.SDesc = sdesc;
                    newItem.SValue = _value;
                    newItem.IMonth = 0;
                    newItem.IYear = 0;
                    newItem.IStatus = 1;
                    newItem.IsNew = true;
                    newItem.Save();
                    DataRow newrow = globalVariables.gv_dtSysparams.NewRow();
                    newrow[SysSystemParameter.FpSBranchIDColumn.ColumnName] = globalVariables.Branch_ID;
                    newrow[SysSystemParameter.SNameColumn.ColumnName] = ParamName;
                    newrow[SysSystemParameter.SDescColumn.ColumnName] = sdesc;
                    newrow[SysSystemParameter.SValueColumn.ColumnName] = _value;
                    newrow[SysSystemParameter.IYearColumn.ColumnName] = 0;
                    newrow[SysSystemParameter.IMonthColumn.ColumnName] = 0;
                    newrow[SysSystemParameter.IStatusColumn.ColumnName] = 1;
                    globalVariables.gv_dtSysparams.Rows.Add(newrow);
                    globalVariables.gv_dtSysparams.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg("Lỗi khi cập nhật giá trị tham số hệ thống:\n" + ex.Message);
            }
        }
        public static string Laygiatrithamsohethong(string ParamName, string defaultval, bool fromDB)
        {
            try
            {
                fromDB = true;
                string reval = defaultval;
                if (fromDB)
                {
                    SqlQuery sqlQuery =
                        new Select().From(SysSystemParameter.Schema).Where(SysSystemParameter.Columns.SName).IsEqualTo(
                            ParamName);
                    SysSystemParameter objSystemParameter = sqlQuery.ExecuteSingle<SysSystemParameter>();
                    if (objSystemParameter != null) reval = objSystemParameter.SValue;
                }
                else
                {
                    DataRow[] arrDR = globalVariables.gv_dtSysparams.Select(SysSystemParameter.SNameColumn.ColumnName + " ='" + ParamName + "'");
                    if (arrDR.Length > 0) reval = Utility.sDbnull(arrDR[0][SysSystemParameter.SValueColumn.ColumnName]);
                }
                return reval;
            }
            catch
            {
                return defaultval;
            }
        }
        public static DataTable LayDulieuDanhmucChung(List<string> lstLoai, bool fromDB)
        {
            try
            {
                DataTable m_NN = new DataTable();
                if (fromDB)
                {
                    m_NN =
                        new Select().From(DmucChung.Schema)
                            .Where(DmucChung.Columns.Loai).In(lstLoai)
                            .OrderAsc(DmucChung.Columns.SttHthi)
                            .ExecuteDataSet().Tables[0];
                }
                else
                {

                    var q = from p in globalVariables.gv_dtDmucChung.AsEnumerable()
                            where lstLoai.Contains(p.Field<string>(DmucChung.Columns.Loai))
                            select p;
                    if (q.Count() <= 0)
                        m_NN =
                        new Select().From(DmucChung.Schema)
                            .Where(DmucChung.Columns.Loai).In("-1")
                            .ExecuteDataSet().Tables[0];
                    else
                        m_NN = q.CopyToDataTable();
                }
                return m_NN;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static DataTable LaydanhsachDoituongKcb()
        {
            try
            {
                DataTable dtData = new DataTable();

                dtData =
                        new Select().From(DmucDoituongkcb.Schema)
                            .OrderAsc(DmucDoituongkcb.Columns.SttHthi)
                            .ExecuteDataSet().Tables[0];

                return dtData;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static DataTable LaydanhsachDoituongKcb(List<string>lstDoituong)
        {
            try
            {
                DataTable dtData = new DataTable();

                dtData =
                        new Select().From(DmucDoituongkcb.Schema)
                        .Where(DmucDoituongkcb.Columns.MaDoituongKcb).In(lstDoituong)
                            .OrderAsc(DmucDoituongkcb.Columns.SttHthi)
                            .ExecuteDataSet().Tables[0];

                return dtData;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static DataTable LaydanhsachBacsi(int departmentID,int noitru)
        {
            try
            {
                DataTable dtData = new DataTable();

                dtData = SPs.DmucLaydanhsachBacsi(departmentID, noitru).GetDataSet().Tables[0];
                return dtData;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
      
        public static DataTable LaydanhsachThunganvien()
        {
            try
            {
                DataTable dtData = new DataTable();

                dtData = SPs.DmucLaydanhsachTnv().GetDataSet().Tables[0];
                return dtData;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static DataTable Laydanhsachnhanvien(string ma_loainv)
        {
            try
            {
                DataTable dtData = new DataTable();

                dtData = SPs.DmucLaydanhsachNhanvien(ma_loainv).GetDataSet().Tables[0];
                return dtData;
            }
            catch (Exception ex)
            {

                return null;
            }
        }
       
        
        public static DmucChung LaydoituongDmucChung(string Loai, string MA)
        {
            try
            {

                return new Select().From(DmucChung.Schema)
                    .Where(DmucChung.Columns.Loai).IsEqualTo(Loai)
                    .And(DmucChung.Columns.Ma).IsEqualTo(MA)
                    .ExecuteSingle<DmucChung>();

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static DataTable LayDulieuDanhmucChung(string Loaidmuc, bool fromDB)
        {
            try
            {
                DataTable m_NN = new DataTable();
                if (fromDB)
                {
                    m_NN =
                        new Select().From(DmucChung.Schema)
                            .Where(DmucChung.Columns.Loai).IsEqualTo(Loaidmuc)
                            .OrderAsc(DmucChung.Columns.SttHthi)
                            .ExecuteDataSet().Tables[0];
                }
                else
                {

                    var q = from p in globalVariables.gv_dtDmucChung.AsEnumerable()
                            where p.Field<string>(DmucChung.Columns.Loai) == Loaidmuc
                            select p;
                    if (q.Count() <= 0)
                        m_NN =
                        new Select().From(DmucChung.Schema)
                            .Where(DmucChung.Columns.Loai).In("-1")
                            .ExecuteDataSet().Tables[0];
                    else
                        m_NN = q.CopyToDataTable();
                }
                return m_NN;
            }
            catch(Exception ex)
            {
                
                return null;
            }
        }
       
        public static DataTable LaydanhsachKhoanoitruTheoBacsi(string username, byte isAdmin,byte noitru)
        {
            DataTable dtData = SPs.DmucLaydanhsachCackhoaKCBtheoBacsi(username, isAdmin, noitru).GetDataSet().Tables[0];
            DataTable dtRevalue = dtData.Clone();
            //NOITRU_NAPKHOANOITRU_THEOKHOADANGNHAP: 1= Khi bác sĩ liên khoa đăng nhập thì tại các chức năng phần nội trú sẽ chỉ hiển thị duy nhất khoa là khoa đăng nhập
            //0= Các chức năng nội trú load tất cả các liên khoa của BS để bác sĩ tự chọn và tìm kiếm BN
            if (!Utility.Byte2Bool(isAdmin) && THU_VIEN_CHUNG.Laygiatrithamsohethong("NOITRU_NAPKHOANOITRU_THEOKHOADANGNHAP", "0", true) == "1")
            {
                DataRow[] arrDr = dtData.Select(DmucKhoaphong.Columns.IdKhoaphong + "=" + globalVariables.idKhoatheoMay.ToString());
                if (arrDr.Length > 0)
                    dtRevalue = arrDr.CopyToDataTable();
            }
            else
                dtRevalue = dtData.Copy();
            return dtRevalue;
        }
        public static DataTable LaydanhsachKhoaKhidangnhap(string username, byte isAdmin)
        {
            DataTable dtData = SPs.DmucLaydanhsachCackhoaKCBtheoBacsi(username, isAdmin, (byte)2).GetDataSet().Tables[0];
            DataTable dtRevalue = dtData.Clone();
            return dtData;
        }
        public static DataTable DmucLaydanhsachCacphongkhamTheoBacsi(string UserName, short? Idkhoa, byte? IsAdmin, byte? Noitru)
        {
            return SPs.DmucLaydanhsachCacphongkhamTheoBacsi(UserName, Idkhoa, IsAdmin, Noitru).GetDataSet().Tables[0];
        }
        public static DateTime GetSysDateTime()
        {
            try
            {
                DataTable dataTable = new DataTable();
                DateTime dt = new DateTime();
                dt = DateTime.Now;
                DateTime dateTime = new SubSonic.InlineQuery().ExecuteScalar<DateTime>("select getdate()");

                return dateTime;
            }
            catch
            {
                return DateTime.Now;
            }
        }
       
        public static int LayIDPhongbanTheoUser(string sUserName)
        {
            int vDepartment_Id = -1;
            try
            {
                SqlQuery sqlQuery = new Select().From(DmucNhanvien.Schema).Where(DmucNhanvien.Columns.UserName).IsEqualTo(sUserName);
                DmucNhanvien objStaff = sqlQuery.ExecuteSingle<DmucNhanvien>();
                if (objStaff != null)
                {
                    vDepartment_Id =Utility.Int32Dbnull( objStaff.IdKhoa,-1);
                    globalVariables.gv_intIDNhanvien = objStaff.IdNhanvien;
                    globalVariables.gv_strTenNhanvien = Utility.sDbnull(objStaff.TenNhanvien);
                    globalVariables.gv_intIDNhanvien = Utility.Int16Dbnull(objStaff.IdNhanvien);
                    globalVariables.gv_intPhongNhanvien = Utility.Int16Dbnull(objStaff.IdPhong);
                }
            }
            catch (Exception ex)
            {
                vDepartment_Id = -1;

            }



            return vDepartment_Id;
        }
        public static void Sapxepthutuin(ref DataTable dataTable,bool BHYT)
        {
            Utility.AddColumToDataTable(ref dataTable, "stt_in", typeof(int));
            List<DmucChung> lst =
                new Select().From(DmucChung.Schema)
                    .Where(DmucChung.Columns.Loai)
                    .IsEqualTo(BHYT ? THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_STT_INPHOI", false) : THU_VIEN_CHUNG.Laygiatrithamsohethong("DICHVU_STT_IN", false))
                    .ExecuteAsCollection<DmucChungCollection>().ToList<DmucChung>();
           
            foreach (DataRow dr in dataTable.Rows)
            {
                var objDmucChung = from p in lst
                                         where p.Ma == Utility.sDbnull(dr["id_loaithanhtoan"])
                                         select p;
                if (objDmucChung != null && objDmucChung.Any())
                {
                    dr["stt_in"] = Utility.Int32Dbnull( objDmucChung.FirstOrDefault().SttHthi);
                    dr["id_loaithanhtoan"] = Utility.Int32Dbnull(objDmucChung.FirstOrDefault().Ma);
                    dr["ten_loaithanhtoan"] =Utility.sDbnull(objDmucChung.FirstOrDefault().Ten);
                    
                }
            }
            int max = dataTable.AsEnumerable().Select(c => c.Field<int>("stt_in")).Max();
            for (int i = 1; i <= max; i++)
            {
                if (dataTable.Select("stt_in=" + i).Length > 0) continue;
                var q = from p in dataTable.AsEnumerable()
                        where Utility.Int32Dbnull(p["stt_in"]) > i
                        select p;
                if (q.Any())
                {
                    int min = q.Select(c => c.Field<int>("stt_in")).Min();
                    if (min != i)
                    {
                        DataRow[] arrDr = dataTable.Select("stt_in=" + min);
                        foreach (DataRow dr in arrDr)
                            dr["stt_in"] = i;
                    }
                }
            }
            foreach (DataRow dr in dataTable.Rows)
            {
                dr["ten_loaithanhtoan"] = dr["stt_in"] + ". " + dr["ten_loaithanhtoan"];
                if (dr["id_loaithanhtoan"].ToString() == "3")
                    dr["ten_loaidichvu"] = dr["stt_in"] + ".1 Trong danh mục BHYT ";
            }
            dataTable.AcceptChanges();
        }
       
        public static int LayIDPhongbanTheoMay(string Code)
        {
            int vDepartment_Id = -1;
            try
            {
                DmucKhoaphong _DmucKhoaphong = new Select().From(DmucKhoaphong.Schema).Where(DmucKhoaphong.Columns.MaKhoaphong).IsEqualTo(Code).ExecuteSingle<DmucKhoaphong>();

                if (_DmucKhoaphong != null)
                {
                    return _DmucKhoaphong.IdKhoaphong;
                }
                return globalVariables.IdKhoaNhanvien;
            }
            catch (Exception ex)
            {
               return globalVariables.IdKhoaNhanvien;

            }
            return vDepartment_Id;
        }
        public static void LoadThamSoHeThong()
        {
            globalVariables.FORMTITLE=THU_VIEN_CHUNG.Laygiatrithamsohethong("FORMTITLE", false);
            globalVariables.LUONGCOBAN = Utility.DecimaltoDbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_LUONGCOBAN", "83000", false), 83000);
            globalVariables.gv_strNoiDKKCBBD = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_NOIDANGKY_KCBBD", "016", false);
            globalVariables.gv_strDiadiem = THU_VIEN_CHUNG.Laygiatrithamsohethong("DIA_DIEM", "Hà Nội", false);
            globalVariables.gv_strNoicapBHYT = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_NOICAP_BHYT", "01", false);

            globalVariables.gv_intChophepchongiathuoc =  Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("CHONGIATHUOC", "0",false),0);
            globalVariables.gv_blnApdungChedoDuyetBHYT = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_TUDONGDUYET", "1", false) == "1";

            globalVariables.gv_GiathuoctheoGiatrongKho = THU_VIEN_CHUNG.Laygiatrithamsohethong("GIATHUOCKHO", "1", false) == "1";
            globalVariables.ChophepNhapkhoLe = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("ChophepNhapkhoLe", "0", false));
            globalVariables.gv_strTuyenBHYT = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_TUYEN", "TW", false);
            globalVariables.TrongGio = THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_TRONGGIO", "0:00-23:59", false);
            globalVariables.TrongNgay = THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_TRONGNGAY", "2,3,4,5,6,7,CN", false);
            globalVariables.gv_intKT_TT_ChuyenCLS_DV = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("KT_TT_ChuyenCLS_DV", "0", false), 0);
            globalVariables.gv_strBHYT_MAQUYENLOI_UUTIEN = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_MAQUYENLOI_UUTIEN", "", false);
            globalVariables.gv_intKT_TT_ChuyenCLS_BHYT = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("KT_TT_ChuyenCLS_BHYT", "0", false), 0);
            globalVariables.gv_strICD_BENH_AN_NGOAI_TRU = THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_ICD_BENH_AN_NGOAI_TRU", "", false);

            globalVariables.gv_intSO_BENH_AN_BATDAU = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_SO_BENH_AN", "-1", false), -1);
            globalVariables.gv_strMA_BHYT_KT = THU_VIEN_CHUNG.Laygiatrithamsohethong("MA_BHYT_KT", "", false);
            globalVariables.gv_strMaQuyenLoiHuongBHYT100Phantram = THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_MAQUYENLOI_HUONG100PHANTRAM", "1,2", false);
            globalVariables.gv_intCHARACTERCASING = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_CHARACTERCASING", "0", false), 0);
            globalVariables.gv_intKIEMTRAMATHEBHYT = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_KIEMTRAMATHE", "0", false), 0);
            globalVariables.gv_intBHYT_TUDONGCHECKTRAITUYEN = Utility.Int32Dbnull(THU_VIEN_CHUNG.Laygiatrithamsohethong("BHYT_TUDONGCHECKTRAITUYEN", "0", false), 0);
            globalVariables.gv_strBOTENDIACHINH = THU_VIEN_CHUNG.Laygiatrithamsohethong("BOTENDIACHINH", "", false);
            if (globalVariablesPrivate.objNhanvien != null)
            {
                globalVariables.gv_strTenNhanvien = globalVariablesPrivate.objNhanvien.TenNhanvien;
                globalVariables.gv_intIDNhanvien = globalVariablesPrivate.objNhanvien.IdNhanvien;
            }
            else
            {
                globalVariables.gv_strTenNhanvien = globalVariables.UserName;
                globalVariables.gv_intIDNhanvien = -1;
            }
            globalVariables.gv_dtQuyenNhanvien = new Select().From(QheNhanvienQuyensudung.Schema).Where(QheNhanvienQuyensudung.Columns.IdNhanvien).IsEqualTo( globalVariables.gv_intIDNhanvien).ExecuteDataSet().Tables[0];
            globalVariables.gv_dtQuyenNhanvien_Dmuc = new Select().From(QheNhanvienDanhmuc.Schema).Where(QheNhanvienDanhmuc.Columns.IdNhanvien).IsEqualTo(globalVariables.gv_intIDNhanvien).ExecuteDataSet().Tables[0];
        }
        /// <summary>
        /// Hàm thực hiện lấy về IpAddress của máy đang login
        /// </summary>
        /// <returns></returns>
        public static string GetIP4Address()
        {
            try
            {
                if (string.IsNullOrEmpty(globalVariables.gv_strIPAddress))
                {
                    string IP4Address = String.Empty;

                    foreach (IPAddress IPA in Dns.GetHostAddresses(Dns.GetHostName()))
                    {
                        if (IPA.AddressFamily == AddressFamily.InterNetwork)
                        {
                            IP4Address = IPA.ToString();
                            break;
                        }
                    }
                    globalVariables.gv_strIPAddress = IP4Address;
                }


                return globalVariables.gv_strIPAddress;
            }
            catch
            { return "NO-IP"; }
        }
         //<summary>
         //hàm thực hiện việc lấy thông tin của địa chỉ mac cho máy tính
         //</summary>
         //<returns></returns>

        public static string GetMACAddress()
        {
            try
            {
                if (string.IsNullOrEmpty(globalVariables.gv_strMacAddress))
                {
                    NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
                    String sMacAddress = string.Empty;
                    foreach (NetworkInterface adapter in nics)
                    {
                        if (sMacAddress == String.Empty)// only return MAC Address from first card  
                        {
                            IPInterfaceProperties properties = adapter.GetIPProperties();
                            sMacAddress = adapter.GetPhysicalAddress().ToString();
                            globalVariables.gv_strMacAddress = sMacAddress;
                        }
                    }
                }
                //  Utility.sDbnull()
                return globalVariables.gv_strMacAddress;
            }
            catch
            { return "NO-ADDRESS"; }
        }
        public static DataTable Laydanhsachnhanvienthuockhoa(int id_khoa)
        {
            DataTable dataTable = new DataTable();
            SqlQuery sqlQuery = new Select().From(DmucNhanvien.Schema)
                .Where(DmucNhanvien.Columns.IdKhoa).IsEqualTo(id_khoa);
            dataTable = sqlQuery.ExecuteDataSet().Tables[0];
            return dataTable;

        }
        public static DataTable Laydanhsachphongthuockhoa(int id_khoa, int PhongChucnang)
        {
            DataTable dataTable = new DataTable();
            SqlQuery sqlQuery = new Select().From(VDmucKhoaphong.Schema)
                .Where(VDmucKhoaphong.Columns.MaCha).IsEqualTo(id_khoa);
            if (PhongChucnang > -1)
                sqlQuery.And(VDmucKhoaphong.Columns.PhongChucnang).IsEqualTo(PhongChucnang);
            sqlQuery.And(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("PHONG");
            dataTable = sqlQuery.ExecuteDataSet().Tables[0];
            return dataTable;

        }
        public static DataTable Laydanhsachphongthuockhoa(string ma_khoa, int PhongChucnang)
        {
            int id_khoa=-1;
            DataTable dataTable = new DataTable();
            DmucKhoaphong _item = new Select().From(DmucKhoaphong.Schema).Where(DmucKhoaphong.Columns.MaKhoaphong).IsEqualTo(ma_khoa).ExecuteSingle<DmucKhoaphong>();
            if (_item != null) id_khoa = _item.IdKhoaphong;
            SqlQuery sqlQuery = new Select().From(VDmucKhoaphong.Schema)
                .Where(VDmucKhoaphong.Columns.MaCha).IsEqualTo(id_khoa);
            if (PhongChucnang > -1)
                sqlQuery.And(VDmucKhoaphong.Columns.PhongChucnang).IsEqualTo(PhongChucnang);
            sqlQuery.And(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("PHONG");
            dataTable = sqlQuery.ExecuteDataSet().Tables[0];
            return dataTable;

        }
        
        public static DataTable Laydanhmuckhoa(string NoitruNgoaitru,int PhongChucnang)
        {
            try
            {
                SqlQuery sqlQuery = new Select().From(VDmucKhoaphong.Schema).Where(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("KHOA");
                if (PhongChucnang > -1)
                    sqlQuery.And(VDmucKhoaphong.Columns.PhongChucnang).IsEqualTo(PhongChucnang);
                if (NoitruNgoaitru !="ALL")
                    sqlQuery.And(VDmucKhoaphong.Columns.NoitruNgoaitru).IsEqualTo(NoitruNgoaitru);
                return sqlQuery.ExecuteDataSet().Tables[0];
            }
            catch
            {
                return null;
            }
        }

        public static DataTable Laydanhmuckhoa(string NoitruNgoaitru, int PhongChucnang,int idkhoa_loaibo)
        {
            try
            {
                SqlQuery sqlQuery = new Select().From(VDmucKhoaphong.Schema).Where(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("KHOA");
                if (PhongChucnang > -1)
                    sqlQuery.And(VDmucKhoaphong.Columns.PhongChucnang).IsEqualTo(PhongChucnang);
                if (idkhoa_loaibo > -1)
                    sqlQuery.And(VDmucKhoaphong.Columns.IdKhoaphong).IsNotEqualTo(idkhoa_loaibo);
                if (NoitruNgoaitru != "ALL")
                    sqlQuery.And(VDmucKhoaphong.Columns.NoitruNgoaitru).IsEqualTo(NoitruNgoaitru);
                return sqlQuery.ExecuteDataSet().Tables[0];
            }
            catch
            {
                return null;
            }
        }
        public static DataTable Laydanhmuckhoa( int idkhoa)
        {
            try
            {
                SqlQuery sqlQuery = new Select().From(VDmucKhoaphong.Schema).Where(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("KHOA");
                sqlQuery.And(VDmucKhoaphong.Columns.IdKhoaphong).IsEqualTo(idkhoa);
               
                return sqlQuery.ExecuteDataSet().Tables[0];
            }
            catch
            {
                return null;
            }
        }
        public static DataTable LaydanhmucPhong(int PhongChucnang)
        {
            try
            {
                SqlQuery sqlQuery =new Select().From(VDmucKhoaphong.Schema).Where(VDmucKhoaphong.Columns.KieuKhoaphong).IsEqualTo("PHONG");
                        if (PhongChucnang > -1)
                    sqlQuery.And(VDmucKhoaphong.Columns.PhongChucnang).IsEqualTo(PhongChucnang);
                        return sqlQuery.ExecuteDataSet().Tables[0];
            }
            catch
            {
                return null;
            }
        }
        public static string MaNhapKho(int LoaiPhieu)
        {
            string MaNhapKho = "";
            DataTable dataTable = SPs.ThuocTaomaphieu(LoaiPhieu).GetDataSet().Tables[0];
            if (dataTable.Rows.Count > 0)
            {
                MaNhapKho = Utility.sDbnull(dataTable.Rows[0][0]);
            }
            return MaNhapKho;
        }

        public static string MaTraLaiKho()
        {
            string MaNhapKho = "";
            DataTable dataTable = SPs.TaoMaphieuTraKhoLeVeKhoChan().GetDataSet().Tables[0];
            if (dataTable.Rows.Count > 0)
            {
                MaNhapKho = Utility.sDbnull(dataTable.Rows[0][0]);
            }
            return MaNhapKho;
        }
        /// <summary>
        /// hàm thực hiện việc lấy thông tin mã phiếu xuât kho cho bệnh nhân
        /// </summary>
        /// <returns></returns>
        public static string MaPhieuXuatBN()
        {
            string MaNhapKho = "";
            DataTable dataTable = SPs.TaoMaphieuXuatthuocBN().GetDataSet().Tables[0];
            if (dataTable.Rows.Count > 0)
            {
                MaNhapKho = Utility.sDbnull(dataTable.Rows[0][0]);
            }
            return MaNhapKho;
        }
        public static string BottomCondition()
        {

            return string.Format("Hệ thống quản lý bệnh viện HIS, Phiếu in lúc : {0} in bởi : {1}",
                                 globalVariables.SysDate,
                                 !string.IsNullOrEmpty(globalVariables.gv_strTenNhanvien) ? globalVariables.gv_strTenNhanvien :
                                 globalVariables.UserName);

        }
        public static string TaoTenDonthuoc(string v_PatientCode, int v_Patient_Id)
        {

            string v_Pres_Name = "";
            v_Pres_Name = Utility.sDbnull(SPs.TaoTenDonthuoc(v_PatientCode, v_Patient_Id).ExecuteScalar(), "");
            return v_Pres_Name;
        }
        public static void CreateXML(DataTable dt)
        {
            try
            {
                if (!dt.Columns.Contains("Logo")) dt.Columns.Add(new DataColumn("Logo", typeof(byte[])));
                if (!dt.Columns.Contains("barcode")) dt.Columns.Add(new DataColumn("barcode", typeof(byte[])));
                bool _XML = Laygiatrithamsohethong("XML", "0", false) == "1";
                string _filePath = "newXML.xml";
                if (!_filePath.Contains(@"\")) _filePath = System.Windows.Forms.Application.StartupPath + @"\Xml4Reports\newXML.xml";
                if (_XML)
                {
                    DataTable newDT = dt.Copy();
                    if (newDT.DataSet != null)
                        newDT.DataSet.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                    else
                    {
                        DataSet ds = new DataSet();
                        ds.Tables.Add(newDT);
                        ds.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg(ex.Message);
            }
        }
        public static void CreateXML(DataSet ds)
        {
            try
            {
                if (!ds.Tables[0].Columns.Contains("Logo")) ds.Tables[0].Columns.Add(new DataColumn("Logo", typeof(byte[])));
                if (!ds.Tables[0].Columns.Contains("barcode")) ds.Tables[0].Columns.Add(new DataColumn("barcode", typeof(byte[])));
                bool _XML = Laygiatrithamsohethong("XML", "0", false) == "1";
                string _filePath = "newXML.xml";
                if (!_filePath.Contains(@"\")) _filePath = System.Windows.Forms.Application.StartupPath + @"\Xml4Reports\newXML.xml";
                if (_XML)
                {
                   
                        ds.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg(ex.Message);
            }
        }
        public static void CreateXML(DataTable dt,string xmlfile)
        {
            try
            {
                if (!dt.Columns.Contains("Logo")) dt.Columns.Add(new DataColumn("Logo",typeof(byte[])));
                if (!dt.Columns.Contains("barcode")) dt.Columns.Add(new DataColumn("barcode", typeof(byte[])));
                bool _XML = Laygiatrithamsohethong("XML", "0", false) == "1";
                string _filePath = xmlfile;
                if (!_filePath.ToUpper().Contains(".XML")) _filePath += ".xml";

                if (!_filePath.Contains(@"\")) _filePath = System.Windows.Forms.Application.StartupPath + @"\Xml4Reports\" + _filePath;
                Utility.CreateFolder(_filePath);
                if (_XML)
                {
                    DataTable newDT = dt.Copy();
                    if (newDT.DataSet != null)
                        newDT.DataSet.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                    else
                    {
                        DataSet ds = new DataSet();
                        ds.Tables.Add(newDT);
                        ds.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                    }
                }
            }
            catch(Exception ex)
            {
                Utility.ShowMsg(ex.Message);
            }
        }
        public static void CreateXML(DataSet ds, string xmlfile)
        {
            try
            {
                if (!ds.Tables[0].Columns.Contains("Logo")) ds.Tables[0].Columns.Add(new DataColumn("Logo", typeof(byte[])));
                if (!ds.Tables[0].Columns.Contains("barcode")) ds.Tables[0].Columns.Add(new DataColumn("barcode", typeof(byte[])));
                bool _XML = Laygiatrithamsohethong("XML", "0", false) == "1";
                string _filePath = xmlfile;
                if (!_filePath.ToUpper().Contains(".XML")) _filePath += ".xml";
                if (!_filePath.Contains(@"\")) _filePath = System.Windows.Forms.Application.StartupPath + @"\Xml4Reports\" + _filePath;
                Utility.CreateFolder(_filePath);
                if (_XML)
                {

                    ds.WriteXml(_filePath, XmlWriteMode.WriteSchema);
                }
            }
            catch (Exception ex)
            {
                Utility.ShowMsg(ex.Message);
            }
        }
       
       
        public static decimal TinhBhytChitra(decimal PhanTramBH, decimal Origin_Price)
        {
            return PhanTramBH * Origin_Price / 100;
        }
        public static decimal TinhBhytChitra(decimal PhanTramBH, decimal Origin_Price, int IsPayment)
        {
            if (IsPayment == 0)
                return PhanTramBH * Origin_Price / 100;
            else
            {
                return 0;
            }
        }
        public static decimal TinhBnhanChitra(decimal PhanTramBH, decimal Origin_Price)
        {
            return (100 - PhanTramBH) * Origin_Price / 100;
        }
        public static decimal TinhBnhanChitra(decimal PhanTramBH, decimal Origin_Price, int IsPayment)
        {
            if (IsPayment == 0)
                return (100 - PhanTramBH) * Origin_Price / 100;
            else
            {
                return Origin_Price;
            }
        }
       
        public static bool IsNgoaiGio()
        {
            try
            {
                GetTrongNgayTrongGio();
                //Kiểm tra ngày hiện tại có trong tham biến không
                if (KT_TRONGNGAY())
                {
                    // Nếu có trong ngày kiểm tra giờ hiện tại có trong giờ ko
                    if (!Utility.IsBetweenManyTimeranges(globalVariables.SysDate, globalVariables.TrongGio))
                    {
                        //Nếu giờ hiện tại không trong giờ tham biến trả về true. Ngoài giờ khám
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return true;

            }
            catch (Exception)
            {
                return false;
            }
        }
        public static void GetTrongNgayTrongGio()
        {
            globalVariables.TrongGio = THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_TRONGGIO", "0:00-23:59", false);
            globalVariables.TrongNgay = THU_VIEN_CHUNG.Laygiatrithamsohethong("KCB_TRONGNGAY", "2,3,4,5,6,7,CN", false);
        }
        
        /// <summary>
        /// Kiểm tra so sánh ngày hiện tại với các ngày trong biến TRONGNGAY
        /// </summary>
        /// <returns></returns>
        static bool KT_TRONGNGAY()
        {
            try
            {
                string[] TrongNgay = globalVariables.TrongNgay.Split(',');
                if (TrongNgay.Length > 0)
                {
                    //So sánh giá trị từng ngày trong mảng.
                    foreach (string s in TrongNgay)
                    {
                        switch (s)
                        {
                            //Thứ 2 : giá trị so sánh = 1;
                            case "2":
                                //Nếu so sánh ngày bằng nhau thì trả về true
                                if (_SoSanhNgay(1))
                                {
                                    return true;
                                }
                                break;
                            //Thứ 3 : giá trị so sánh = 2;
                            case "3":
                                if (_SoSanhNgay(2))
                                {
                                    return true;
                                }
                                break;
                            //Thứ 4 : giá trị so sánh = 3;
                            case "4":
                                if (_SoSanhNgay(3))
                                {
                                    return true;
                                }
                                break;
                            //Thứ 5 : giá trị so sánh = 4;
                            case "5":
                                if (_SoSanhNgay(4))
                                {
                                    return true;
                                }
                                break;
                            //Thứ 6 : giá trị so sánh = 5;
                            case "6":
                                if (_SoSanhNgay(5))
                                {
                                    return true;
                                }
                                break;
                            //Thứ 7 : giá trị so sánh = 6;
                            case "7":
                                if (_SoSanhNgay(6))
                                {
                                    return true;
                                }
                                break;
                            //Thứ CN : giá trị so sánh = 0;
                            case "CN":
                                if (_SoSanhNgay(0))
                                {
                                    return true;
                                }
                                break;
                        }
                    }
                    //Nếu hết các giá trị trong mảng ko có giá trị nào bằng ngày hiện tại thì trả về false
                    return false;
                }
                //Nếu mảng giá trị nhỏ hơn không là ko có tham biến thì trả về true
                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }
        /// <summary>
        /// KIểm tra so sánh ngày trong biến truyền vào với ngày hiện tại. Nếu bằng nhau thì trả về true else false
        /// </summary>
        /// <param name="Ngay"></param>
        /// <returns></returns>
        static bool _SoSanhNgay(int Ngay)
        {
            try
            {
                return (int)GetSysDateTime().DayOfWeek == Ngay;
            }
            catch (Exception)
            {
                return false;
            }
        }
       
      
        public static string GetThamSo_ThuTu()
        {
            string thamso = "THANG";
            SqlQuery sqlQuery =
                new Select().From(SysSystemParameter.Schema).Where(SysSystemParameter.Columns.SName).IsEqualTo(
                    "STT_KHAM");
            SysSystemParameter objSystemParameter = sqlQuery.ExecuteSingle<SysSystemParameter>();
            if (objSystemParameter != null) thamso = objSystemParameter.SValue;
            return thamso;
        }
        public static string GetMaPhieuThu(DateTime dateTime, int LoaiPhieu)
        {
            
            return Utility.sDbnull(SPs.TaoMaPhieuthu(dateTime, LoaiPhieu).ExecuteScalar<string>(), "");
        }
       
       public static string  LayMaDviLamViec()
       {
           return "NO";
       }
       public static bool IsBaoHiem(byte IdLoaidoituongKcb)
       {
           return IdLoaidoituongKcb == (byte)0;
       }
       public static bool IsBaoHiem(byte? IdLoaidoituongKcb)
       {
           return Utility.ByteDbnull(IdLoaidoituongKcb, 1) == (byte)0;
       }
       
      
       public static string GetThanhToan_TraiTuyen()
       {
           string sPaymentFlow = "";
           SqlQuery sqlQuery = new Select().From(SysSystemParameter.Schema)
               .Where(SysSystemParameter.Columns.SName).IsEqualTo("TRAITUYEN");
           SysSystemParameter objSystemParameter = sqlQuery.ExecuteSingle<SysSystemParameter>();
           if (objSystemParameter != null) sPaymentFlow = objSystemParameter.SValue;
           return sPaymentFlow;
       }
       
       public static string MaKieuThanhToan(int PaymentType_ID)
       {
           string MaKieu = "";
           switch (PaymentType_ID)
           {
               case 0:
                   MaKieu = "PHI_DVYC";
                   break;
               case 1:
                   MaKieu = "KHAM";
                   break;
               case 2:
                   MaKieu = "CLS";
                   break;
               case 3:
                   MaKieu = "THUOC";
                   break;
               case 4:
                   MaKieu = "GIUONG";
                   break;
               case 5:
                   MaKieu = "VT";
                   break;
               case 6:
                   MaKieu = "TAMUNG";
                   break;
               case 7:
                   MaKieu = "PHIEU_AN";
                   break;
               case 8:
                   MaKieu = "GOIDV";
                   break;
               case 9:
                   MaKieu = "CPTHEM";
                   break;
               case 10:
                   MaKieu = "SO_KHAM";
                   break;
               case 11:
                   MaKieu = "CONG_TIEM";
                   break;
           }
           return MaKieu;
       }
      
        
    }
   
}
