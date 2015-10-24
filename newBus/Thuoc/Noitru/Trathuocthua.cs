using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Transactions;
using VNS.HIS.DAL;
using VNS.Libs;
using SubSonic;
using NLog;
using VNS.Properties;
namespace VNS.HIS.NGHIEPVU.THUOC
{

    public class Trathuocthua
    {
        private NLog.Logger log;
        public Trathuocthua()
        {
            log = NLog.LogManager.GetCurrentClassLogger();
        }

        public static bool ThuocNoitruKiemtraThuoctralai(long id_capphat, long id_donthuoc)
        {
            return SPs.ThuocNoitruKiemtraThuoctralai(id_capphat, id_donthuoc).GetDataSet().Tables[0].Rows.Count > 0;
        }
        public ActionResult CappnhatPhieuTrathuocthua(TPhieutrathuocthua _phieutra, List<long> lstIdCt)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {

                        new Update(TPhieutrathuocthua.Schema).Set(TPhieutrathuocthua.NgayLapphieuColumn).EqualTo(_phieutra.NgayLapphieu.Date)
                            .Set(TPhieutrathuocthua.NguoiLapphieuColumn).EqualTo(_phieutra.NguoiLapphieu)
                            .Set(TPhieutrathuocthua.IdKhoatraColumn).EqualTo(_phieutra.IdKhoatra)
                            .Set(TPhieutrathuocthua.IdKhonhanColumn).EqualTo(_phieutra.IdKhonhan)
                             .Set(TPhieutrathuocthua.NgaySuaColumn).EqualTo(_phieutra.NgaySua.Value)
                            .Set(TPhieutrathuocthua.NguoiSuaColumn).EqualTo(_phieutra.NguoiSua)
                            .Where(TPhieutrathuocthua.IdColumn).IsEqualTo(_phieutra.Id).Execute();

                        new Update(TPhieuCapphatChitiet.Schema)
                           .Set(TPhieuCapphatChitiet.Columns.IdPhieutralai).EqualTo(-1)
                            .Set(TPhieuCapphatChitiet.Columns.TrangthaiTralai).EqualTo(0)
                           .Where(TPhieuCapphatChitiet.Columns.IdPhieutralai).IsEqualTo(_phieutra.Id)
                           .Execute();

                        new Update(TPhieuCapphatChitiet.Schema)
                            .Set(TPhieuCapphatChitiet.Columns.IdPhieutralai).EqualTo(_phieutra.Id)
                             .Set(TPhieuCapphatChitiet.Columns.TrangthaiTralai).EqualTo(1)
                            .Where(TPhieuCapphatChitiet.Columns.IdChitiet).In(lstIdCt)
                            .Execute();

                        Scope.Complete();
                    }
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi cập nhật phiếu trả thuốc thừa", ex);
                return ActionResult.Error;
            }
        }

        public ActionResult ThemPhieutrathuocthua(TPhieutrathuocthua _phieutra, List<long> lstIdCt)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        _phieutra.IsNew = true;
                        _phieutra.Save();
                        if (_phieutra.Id > 0)
                        {
                            new Update(TPhieuCapphatChitiet.Schema)
                            .Set(TPhieuCapphatChitiet.Columns.IdPhieutralai).EqualTo(_phieutra.Id)
                            .Set(TPhieuCapphatChitiet.Columns.TrangthaiTralai).EqualTo(1)
                            .Where(TPhieuCapphatChitiet.Columns.IdChitiet).In(lstIdCt)
                            .Execute();
                        }

                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi thêm phiếu trả thuốc thừa", ex);
                return ActionResult.Error;
            }
        }
        public ActionResult XoaPhieuTrathuocthua(int idphieu)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        new Delete().From(TPhieutrathuocthua.Schema).Where(TPhieutrathuocthua.Columns.Id).IsEqualTo(idphieu).Execute();
                        new Update(TPhieuCapphatChitiet.Schema)
                          .Set(TPhieuCapphatChitiet.Columns.IdPhieutralai).EqualTo(-1)
                          .Set(TPhieuCapphatChitiet.Columns.TrangthaiTralai).EqualTo(0)
                          .Where(TPhieuCapphatChitiet.Columns.IdPhieutralai).IsEqualTo(idphieu)
                          .Execute();
                        Scope.Complete();
                    }
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi xóa phiếu trả thuốc thừa", ex);
                return ActionResult.Error;
            }
        }
      
        public ActionResult Kiemtratonthuoc(TPhieuCapphatChitietCollection lstChitiet, short ID_KHO)
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                foreach (TPhieuCapphatChitiet pres in lstChitiet)
                {
                    if (pres.DaLinh == 1)//Chưa được lĩnh mới kiểm tra
                    {
                        _result = Kiemtrasoluongthuoctrongkho(pres, ID_KHO);
                        if (_result != ActionResult.Success) return _result;
                    }
                }
                return ActionResult.Success;
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        ActionResult Kiemtrasoluongthuoctrongkho(TPhieuCapphatChitiet pres, int ID_KHO)
        {

            int id_thuoc = pres.IdThuoc;
            DmucThuoc _drug = new Select().From(DmucThuoc.Schema).Where(DmucThuoc.IdThuocColumn).IsEqualTo(id_thuoc).ExecuteSingle<DmucThuoc>();
            if (_drug == null) return ActionResult.UNKNOW;
            string Drug_name = _drug.TenThuoc;
            int so_luong = pres.SoLuong;
            int SoLuongTon = CommonLoadDuoc.SoLuongTonTrongKho(pres.IdDonthuoc, ID_KHO, id_thuoc, (int)pres.IdThuockho.Value, 1, (byte)1);
            if (SoLuongTon < so_luong)
            {
                Utility.ShowMsg(string.Format("Bạn không thể xác nhận đơn thuốc,Vì thuốc :{0} số lượng tồn hiện tại trong kho({1}) không đủ cấp cho số lượng yêu cầu({2})\n Mời bạn xem lại số lượng", Drug_name, SoLuongTon.ToString(), so_luong.ToString()));
                return ActionResult.NotEnoughDrugInStock;
            }
            return ActionResult.Success;
        }

       

    }
}
