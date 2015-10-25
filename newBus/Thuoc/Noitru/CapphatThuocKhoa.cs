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

    public class CapphatThuocKhoa
    {
        private NLog.Logger log;
        public CapphatThuocKhoa()
        {
            log = NLog.LogManager.GetCurrentClassLogger();
        }
        public ActionResult Kiemtrathuocxacnhan(TPhieuNhapxuatthuoc objPhieuNhap, TPhieuNhapxuatthuocChitiet objPhieuNhapCt, ref string errMsg)
        {
            TThuockhoCollection vCollection = new TThuockhoController().FetchByQuery(
              TThuockho.CreateQuery()
              .WHERE(TThuockho.IdKhoColumn.ColumnName, Comparison.Equals, objPhieuNhap.IdKhoxuat)
              .AND(TThuockho.IdThuocColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.IdThuoc)
              .AND(TThuockho.NgayHethanColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.NgayHethan.Date)
              .AND(TThuockho.GiaNhapColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.GiaNhap)
              .AND(TThuockho.GiaBanColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.GiaBan)
              .AND(TThuockho.MaNhacungcapColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.MaNhacungcap)
              .AND(TThuockho.SoLoColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.SoLo)
               .AND(TThuockho.NgayNhapColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.NgayNhap)
                .AND(TThuockho.GiaBhytColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.GiaBhyt)
              .AND(TThuockho.VatColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.Vat)
              );

            if (vCollection.Count <= 0)
            {
                errMsg = string.Format("ID thuốc={0}, không tồn tại trong kho {1}", objPhieuNhapCt.IdThuoc.ToString(), objPhieuNhap.IdKhonhap.ToString());
                return ActionResult.Exceed;//Lỗi không có dòng dữ liệu trong bảng kho-thuốc
            }
            int SluongChia=Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
            if(SluongChia<=0) SluongChia=1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
            int SoLuong = vCollection[0].SoLuong;
            if (SluongChia > 1)
                SoLuong = vCollection[0].SoLuong * SluongChia;
            if (SoLuong < 0)
            {
                errMsg = string.Format("ID thuốc={0}, Số lượng còn trong kho {1}, Số lượng bị trừ {2}", objPhieuNhapCt.IdThuoc.ToString(), vCollection[0].SoLuong.ToString(), objPhieuNhapCt.SoLuong.ToString());
                return ActionResult.NotEnoughDrugInStock;//Thuốc đã sử dụng nhiều nên không thể hủy
            }
            return ActionResult.Success;
        }
        public ActionResult XacNhanPhieuCapphatThuocBenhnhanNoitru(TPhieuNhapxuatthuoc objPhieuNhap, DateTime ngayxacnhan, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();
                        objPhieuNhap.NgayXacnhan = ngayxacnhan;
                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra đề phòng Kho A-->Xuất kho B. Kho B xác nhận-->Xuất kho C. Kho B hủy xác nhận. Kho C xác nhận dẫn tới việc kho B chưa có thuốc để trừ kho

                            ActionResult _Kiemtrathuocxacnhan = Kiemtrathuocxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuocxacnhan != ActionResult.Success) return _Kiemtrathuocxacnhan;

                            long idthuockho = -1;

                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int _SoLuong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      _SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhonhap,
                                                                      objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, idthuockho, ngayxacnhan, objPhieuNhapCt.GiaBhyt, objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);

                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhoxuat, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          Utility.Int32Dbnull(objPhieuNhapCt.SoLuong), objPhieuNhapCt.IdChuyen,
                                                          objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo,
                                                          objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);

                            sp.Execute();
                            if (idthuockho > 0)
                                new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(idthuockho)
                                   .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            else
                                idthuockho = Utility.Int64Dbnull(objPhieuNhapCt.IdThuockho, -1);
                            objPhieuNhapCt.IdThuockho = idthuockho;
                            //Insert dòng kho nhập
                            TBiendongThuoc objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);

                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);
                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.Noitru = objPhieuNhap.NoiTru;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdThuockho);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhonhap);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;

                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuNhapKho;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuNhapKho);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();
                            //Insert dòng của kho xuất
                            objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);

                            objXuatNhap.Noitru = objPhieuNhap.NoiTru;
                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.IdChuyen = -1;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdChuyen);
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhoxuat);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuXuatKhoBenhNhan;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuXuatKhoBenhNhan);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();

                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(globalVariables.gv_intIDNhanvien)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(globalVariables.UserName)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(ngayxacnhan)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(1)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi xác nhận phiếu cấp phát thuốc-khoa", ex);
                return ActionResult.Error;
            }
        }
        public ActionResult XacNhanPhieuCapphatThuoc(TPhieuNhapxuatthuoc objPhieuNhap, DateTime ngayxacnhan, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();
                        objPhieuNhap.NgayXacnhan = ngayxacnhan;
                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra đề phòng Kho A-->Xuất kho B. Kho B xác nhận-->Xuất kho C. Kho B hủy xác nhận. Kho C xác nhận dẫn tới việc kho B chưa có thuốc để trừ kho

                            ActionResult _Kiemtrathuocxacnhan = Kiemtrathuocxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuocxacnhan != ActionResult.Success) return _Kiemtrathuocxacnhan;

                            long idthuockho = -1;
                            
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int _SoLuong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      _SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhonhap,
                                                                      objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, idthuockho, ngayxacnhan, objPhieuNhapCt.GiaBhyt, objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);
                           
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhoxuat, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          Utility.Int32Dbnull(objPhieuNhapCt.SoLuong), objPhieuNhapCt.IdChuyen,
                                                          objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo,
                                                          objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);

                            sp.Execute();
                            if (idthuockho > 0)
                                new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(idthuockho)
                                   .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            else
                                idthuockho =Utility.Int64Dbnull( objPhieuNhapCt.IdThuockho,-1);
                            objPhieuNhapCt.IdThuockho = idthuockho;
                            //Insert dòng kho nhập
                            TBiendongThuoc objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);

                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);
                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia,1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdThuockho);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhonhap);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;

                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuNhapKho;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuNhapKho);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();
                            //Insert dòng của kho xuất
                            objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);
                            
                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);

                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.IdChuyen = -1;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdChuyen);
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhoxuat);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuXuatKhoa;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuXuatKhoa);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();

                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(globalVariables.gv_intIDNhanvien)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(globalVariables.UserName)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(ngayxacnhan)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(1)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi xác nhận phiếu cấp phát thuốc-khoa", ex);
                return ActionResult.Error;
            }
        }
        public ActionResult XacNhanPhieuTrathuocKholeBNveKhoLeNoitru(TPhieuNhapxuatthuoc objPhieuNhap, DateTime ngayxacnhan, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();
                        objPhieuNhap.NgayXacnhan = ngayxacnhan;
                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra đề phòng Kho A-->Xuất kho B. Kho B xác nhận-->Xuất kho C. Kho B hủy xác nhận. Kho C xác nhận dẫn tới việc kho B chưa có thuốc để trừ kho

                            ActionResult _Kiemtrathuocxacnhan = Kiemtrathuocxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuocxacnhan != ActionResult.Success) return _Kiemtrathuocxacnhan;
                            long idthuockho = -1;
                            //Nhập vào tủ trực
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      objPhieuNhapCt.SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhonhap,
                                                                      objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau,
                                                                      -1, idthuockho, ngayxacnhan, objPhieuNhapCt.GiaBhyt, objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            //Lấy về Id_thuockho tương ứng trong tủ trực
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);

                            //Trừ tủ trực theo đơn vị chia
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int _SoLuong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhoxuat, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          _SoLuong, objPhieuNhapCt.IdChuyen,
                                                          objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo,
                                                          objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);

                            sp.Execute();
                            if (idthuockho > 0)
                                new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(idthuockho)
                                   .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            else
                                idthuockho = Utility.Int64Dbnull(objPhieuNhapCt.IdThuockho, -1);
                            objPhieuNhapCt.IdThuockho = idthuockho;
                            //Insert dòng kho nhập
                            TBiendongThuoc objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);
                            objXuatNhap.GiaBhytCu = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhytCu);

                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);
                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.Noitru = 0;
                            objXuatNhap.QuayThuoc = 0;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdThuockho);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhonhap);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;//Chính là khoa trả

                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();
                            //Insert dòng của kho xuất
                            objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);
                            objXuatNhap.GiaBhytCu = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhytCu);

                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.Noitru = 0;
                            objXuatNhap.QuayThuoc = 0;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdChuyen);
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhoxuat);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();

                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(globalVariables.gv_intIDNhanvien)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(globalVariables.UserName)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(ngayxacnhan)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(1)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi xác nhận phiếu trả thuốc từ tủ trực khoa nội trú về kho lẻ", ex);
                return ActionResult.Error;
            }
        }
        public ActionResult XacNhanPhieuTrathuocTutrucKhoaveKho(TPhieuNhapxuatthuoc objPhieuNhap, DateTime ngayxacnhan, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();
                        objPhieuNhap.NgayXacnhan = ngayxacnhan;
                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra đề phòng Kho A-->Xuất kho B. Kho B xác nhận-->Xuất kho C. Kho B hủy xác nhận. Kho C xác nhận dẫn tới việc kho B chưa có thuốc để trừ kho

                            ActionResult _Kiemtrathuocxacnhan = Kiemtrathuocxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuocxacnhan != ActionResult.Success) return _Kiemtrathuocxacnhan;
                            long idthuockho = -1;
                           //Nhập vào tủ trực
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      objPhieuNhapCt.SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhonhap,
                                                                      objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau,
                                                                      -1, idthuockho, ngayxacnhan, objPhieuNhapCt.GiaBhyt, objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            //Lấy về Id_thuockho tương ứng trong tủ trực
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);
                            
                            //Trừ tủ trực theo đơn vị chia
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int _SoLuong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhoxuat, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          _SoLuong, objPhieuNhapCt.IdChuyen,
                                                          objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo,
                                                          objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);

                            sp.Execute();
                            if (idthuockho > 0)
                                new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(idthuockho)
                                   .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            else
                                idthuockho = Utility.Int64Dbnull(objPhieuNhapCt.IdThuockho, -1);
                            objPhieuNhapCt.IdThuockho = idthuockho;
                            //Insert dòng kho nhập
                            TBiendongThuoc objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);
                            objXuatNhap.GiaBhytCu = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhytCu);

                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);
                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.Noitru = 0;
                            objXuatNhap.QuayThuoc = 0;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdThuockho);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhonhap);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;//Chính là khoa trả

                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuNhapTraKhoLe;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuNhapTraKhoLe);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();
                            //Insert dòng của kho xuất
                            objXuatNhap = new TBiendongThuoc();
                            objXuatNhap.IdPhieu = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieu);
                            objXuatNhap.IdPhieuChitiet = Utility.Int32Dbnull(objPhieuNhapCt.IdPhieuchitiet);
                            objXuatNhap.MaPhieu = Utility.sDbnull(objPhieuNhap.MaPhieu);
                            objXuatNhap.DonGia = Utility.DecimaltoDbnull(objPhieuNhapCt.DonGia);
                            objXuatNhap.NgayNhap = objPhieuNhapCt.NgayNhap;
                            objXuatNhap.GiaBan = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBan);
                            objXuatNhap.GiaNhap = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaNhap);

                            objXuatNhap.GiaPhuthuDungtuyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuDungtuyen);
                            objXuatNhap.GiaPhuthuTraituyen = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaPhuthuTraituyen);
                            objXuatNhap.GiaBhyt = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhyt);
                            objXuatNhap.GiaBhytCu = Utility.DecimaltoDbnull(objPhieuNhapCt.GiaBhytCu);

                            objXuatNhap.SoHoadon = Utility.sDbnull(objPhieuNhap.SoHoadon);
                            objXuatNhap.PhuThu = 0;
                            objXuatNhap.IdChuyen = objPhieuNhapCt.IdChuyen;
                            objXuatNhap.DuTru = objPhieuNhap.DuTru;
                            objXuatNhap.Noitru = 0;
                            objXuatNhap.QuayThuoc = 0;
                            objXuatNhap.SoLuong = Utility.Int32Dbnull(objPhieuNhapCt.SoLuong);
                            objXuatNhap.SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            objXuatNhap.NgayTao = globalVariables.SysDate;
                            objXuatNhap.NguoiTao = globalVariables.UserName;
                            objXuatNhap.IdThuockho = Utility.Int32Dbnull(objPhieuNhapCt.IdChuyen);
                            objXuatNhap.ThanhTien = Utility.DecimaltoDbnull(objPhieuNhapCt.ThanhTien);
                            objXuatNhap.IdThuoc = Utility.Int32Dbnull(objPhieuNhapCt.IdThuoc);
                            objXuatNhap.Vat = Utility.Int32Dbnull(objPhieuNhap.Vat);
                            objXuatNhap.IdNhanvien = Utility.Int16Dbnull(objPhieuNhap.IdNhanvien);
                            objXuatNhap.IdKho = Utility.Int16Dbnull(objPhieuNhap.IdKhoxuat);
                            objXuatNhap.NgayHethan = objPhieuNhapCt.NgayHethan.Date;
                            objXuatNhap.MaNhacungcap = objPhieuNhapCt.MaNhacungcap;
                            objXuatNhap.SoChungtuKemtheo = objPhieuNhap.SoChungtuKemtheo;
                            objXuatNhap.SoLo = objPhieuNhapCt.SoLo;
                            objXuatNhap.SoDky = objPhieuNhapCt.SoDky;
                            objXuatNhap.SoQdinhthau = objPhieuNhapCt.SoQdinhthau;
                            objXuatNhap.MaLoaiphieu = (byte)LoaiPhieu.PhieuNhapTraKhoLe;
                            objXuatNhap.TenLoaiphieu = Utility.TenLoaiPhieu(LoaiPhieu.PhieuNhapTraKhoLe);
                            objXuatNhap.NgayBiendong = objPhieuNhap.NgayXacnhan;
                            objXuatNhap.NgayHoadon = objPhieuNhap.NgayHoadon;
                            objXuatNhap.KieuThuocvattu = objPhieuNhapCt.KieuThuocvattu;
                            objXuatNhap.IdKhoaLinh = objPhieuNhap.IdKhoalinh;
                            objXuatNhap.IsNew = true;
                            objXuatNhap.Save();

                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(globalVariables.gv_intIDNhanvien)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(globalVariables.UserName)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(ngayxacnhan)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(1)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi xác nhận phiếu trả thuốc từ tủ trực khoa nội trú về kho lẻ",ex);
                return ActionResult.Error;
            }
        }
       
        
        public ActionResult Kiemtrathuochuyxacnhan(TPhieuNhapxuatthuoc objPhieuNhap, TPhieuNhapxuatthuocChitiet objPhieuNhapCt, ref string errMsg)
        {
            TThuockhoCollection vCollection = new TThuockhoController().FetchByQuery(
              TThuockho.CreateQuery()
              .WHERE(TThuockho.IdKhoColumn.ColumnName, Comparison.Equals, objPhieuNhap.IdKhonhap)
              .AND(TThuockho.IdThuocColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.IdThuoc)
              .AND(TThuockho.NgayHethanColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.NgayHethan.Date)
              .AND(TThuockho.GiaNhapColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.GiaNhap)
              .AND(TThuockho.GiaBanColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.GiaBan)
              .AND(TThuockho.MaNhacungcapColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.MaNhacungcap)
              .AND(TThuockho.SoLoColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.SoLo)
              .AND(TThuockho.VatColumn.ColumnName, Comparison.Equals, objPhieuNhapCt.Vat)
              );

            if (vCollection.Count <= 0)
            {
                errMsg = string.Format("ID thuốc={0}, không tồn tại trong kho {1}", objPhieuNhapCt.IdThuoc.ToString(), objPhieuNhap.IdKhonhap.ToString());
                return ActionResult.Exceed;//Lỗi không có dòng dữ liệu trong bảng kho-thuốc
            }
            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
            int SoLuong = vCollection[0].SoLuong;
            if (SluongChia > 1)
                SoLuong = vCollection[0].SoLuong * SluongChia;

            SoLuong = SoLuong - objPhieuNhapCt.SoLuong;
            if (SoLuong < 0)
            {
                errMsg = string.Format("ID thuốc={0}, Số lượng còn trong kho {1}, Số lượng bị trừ {2}", objPhieuNhapCt.IdThuoc.ToString(), vCollection[0].SoLuong.ToString(), objPhieuNhapCt.SoLuong.ToString());
                return ActionResult.NotEnoughDrugInStock;//Thuốc đã sử dụng nhiều nên không thể hủy
            }
            return ActionResult.Success;
        }
        public ActionResult HuyXacNhanPhieuCapphatThuoc(TPhieuNhapxuatthuoc objPhieuNhap, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();

                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra ở kho nhập xem thuốc đã sử dụng chưa
                            ActionResult _Kiemtrathuochuyxacnhan = Kiemtrathuochuyxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuochuyxacnhan != ActionResult.Success) return _Kiemtrathuochuyxacnhan;
                            //Xóa biến động kho nhập
                            new Delete().From(TBiendongThuoc.Schema)
                                .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                                .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                                .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuNhapKho).Execute();
                            //Xóa biến động kho xuất
                            new Delete().From(TBiendongThuoc.Schema)
                               .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                               .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                               .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuXuatKhoa).Execute();

                            new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(-1)
                              .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            long id_thuockho = -1;
                            //Nhập lại kho phát
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      objPhieuNhapCt.SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhoxuat, objPhieuNhapCt.MaNhacungcap,
                                                                      objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, id_thuockho, objPhieuNhapCt.NgayNhap, objPhieuNhapCt.GiaBhyt
                                                                      , objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            //Trừ thuốc từ tủ trực
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int SluongChuachia = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhonhap, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          SluongChuachia, objPhieuNhapCt.IdThuockho, objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);
                            sp.Execute();
                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(0)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi hủy xác nhận phiếu cấp phát thuốc-khoa", ex);
                return ActionResult.Error;
            }
        }
        public ActionResult HuyXacNhanPhieuCapphatkholeBNNoitru(TPhieuNhapxuatthuoc objPhieuNhap, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();

                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra ở kho nhập xem thuốc đã sử dụng chưa
                            ActionResult _Kiemtrathuochuyxacnhan = Kiemtrathuochuyxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuochuyxacnhan != ActionResult.Success) return _Kiemtrathuochuyxacnhan;
                            //Xóa biến động kho nhập
                            new Delete().From(TBiendongThuoc.Schema)
                                .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                                .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                                .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuNhapKho).Execute();
                            //Xóa biến động kho xuất
                            new Delete().From(TBiendongThuoc.Schema)
                               .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                               .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                               .And(TBiendongThuoc.Columns.Noitru).IsEqualTo(1)
                               .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuXuatKhoBenhNhan).Execute();

                            new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(-1)
                              .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                            long id_thuockho = -1;
                            //Nhập lại kho phát
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      objPhieuNhapCt.SoLuong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhoxuat, objPhieuNhapCt.MaNhacungcap,
                                                                      objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, id_thuockho, objPhieuNhapCt.NgayNhap, objPhieuNhapCt.GiaBhyt
                                                                      , objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            //Trừ thuốc từ tủ trực
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int SluongChuachia = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhonhap, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          SluongChuachia, objPhieuNhapCt.IdThuockho, objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);
                            sp.Execute();
                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(0)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi hủy xác nhận phiếu cấp phát thuốc-khoa", ex);
                return ActionResult.Error;
            }
        }

        public ActionResult HuyXacNhanPhieuTrathuocTutrucKhoaVeKho(TPhieuNhapxuatthuoc objPhieuNhap, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();

                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra ở kho nhập xem thuốc đã sử dụng chưa
                            ActionResult _Kiemtrathuochuyxacnhan = Kiemtrathuochuyxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuochuyxacnhan != ActionResult.Success) return _Kiemtrathuochuyxacnhan;
                            //Xóa biến động kho nhập
                            new Delete().From(TBiendongThuoc.Schema)
                                .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                                .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                                .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuNhapTraKhoLe).Execute();
                            //Xóa biến động kho xuất
                            new Delete().From(TBiendongThuoc.Schema)
                               .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                               .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                               .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuNhapTraKhoLe).Execute();

                            //Cộng trả lại kho xuất
                            long idthuockho = -1;
                            //Nhập theo tủ trực theo số lượng chia
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int Sluong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      Sluong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhoxuat, objPhieuNhapCt.MaNhacungcap,
                                                                      objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, idthuockho, objPhieuNhapCt.NgayNhap, objPhieuNhapCt.GiaBhyt,
                                                                      objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);
                           //Xuất thứ nguyên từ kho nhận
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhonhap, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          objPhieuNhapCt.SoLuong, objPhieuNhapCt.IdThuockho, objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);
                            sp.Execute();
                            //Cạp nhật lại id_thuockho =-1(giá trị này được update khi xác nhận phiếu). Giá trị id_chuyen cho biết chuyển từ id_thuockho của kho nội trú.
                            new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(-1)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(0)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi hủy xác nhận phiếu trả thuốc từ tủ trực về kho lẻ nội trú",ex);
                return ActionResult.Error;
            }
        }
        public ActionResult HuyXacNhanPhieuTrathuocTuKholeBNNoitruVeKhoLeNoitru(TPhieuNhapxuatthuoc objPhieuNhap, ref string errMsg)
        {
            HisDuocProperties objHisDuocProperties = PropertyLib._HisDuocProperties;
            string errorMessage = "";
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        SqlQuery sqlQuery = new Select().From(TPhieuNhapxuatthuocChitiet.Schema)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu);
                        TPhieuNhapxuatthuocChitietCollection objPhieuNhapCtCollection =
                            sqlQuery.ExecuteAsCollection<TPhieuNhapxuatthuocChitietCollection>();

                        foreach (TPhieuNhapxuatthuocChitiet objPhieuNhapCt in objPhieuNhapCtCollection)
                        {
                            //Kiểm tra ở kho nhập xem thuốc đã sử dụng chưa
                            ActionResult _Kiemtrathuochuyxacnhan = Kiemtrathuochuyxacnhan(objPhieuNhap, objPhieuNhapCt, ref errMsg);
                            if (_Kiemtrathuochuyxacnhan != ActionResult.Success) return _Kiemtrathuochuyxacnhan;
                            //Xóa biến động kho nhập
                            new Delete().From(TBiendongThuoc.Schema)
                                .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                                .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                                .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru).Execute();
                            //Xóa biến động kho xuất
                            new Delete().From(TBiendongThuoc.Schema)
                               .Where(TBiendongThuoc.IdPhieuColumn).IsEqualTo(objPhieuNhap.IdPhieu)
                               .And(TBiendongThuoc.IdPhieuChitietColumn).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet)
                               .And(TBiendongThuoc.MaLoaiphieuColumn).IsEqualTo((byte)LoaiPhieu.PhieuTrathuocTukhoBNNoitruVeKhoLeNoitru).Execute();

                            //Cộng trả lại kho xuất
                            long idthuockho = -1;
                            //Nhập theo tủ trực theo số lượng chia
                            int SluongChia = Utility.Int32Dbnull(objPhieuNhapCt.SluongChia, 1);
                            if (SluongChia <= 0) SluongChia = 1;//Nếu lỗi do người dùng sửa tay thì tự động đặt=1
                            //Chú ý khi lập phiếu xuất thuốc tủ trực thì
                            //objPhieuNhapCt.SoLuong= số lượng đã chia-->Cần trừ số lượng trong kho xuất theo số lượng nguyên gốc. Tức là phải lấy số lượng này / số lượng chia
                            int Sluong = objPhieuNhapCt.SoLuong * SluongChia;//Số lượng thực sự bị mất khỏi kho xuất(khi xuất thuốc sang tủ trực)
                            StoredProcedure sp = SPs.ThuocNhapkhoOutput(objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                                      Sluong, Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                                      objPhieuNhapCt.IdThuoc, objPhieuNhap.IdKhoxuat, objPhieuNhapCt.MaNhacungcap,
                                                                      objPhieuNhapCt.SoLo, objPhieuNhapCt.SoDky, objPhieuNhapCt.SoQdinhthau, -1, idthuockho, objPhieuNhapCt.NgayNhap, objPhieuNhapCt.GiaBhyt,
                                                                      objPhieuNhapCt.GiaPhuthuDungtuyen, objPhieuNhapCt.GiaPhuthuTraituyen, objPhieuNhapCt.KieuThuocvattu);
                            sp.Execute();
                            idthuockho = Utility.Int64Dbnull(sp.OutputValues[0], -1);
                            //Xuất thứ nguyên từ kho nhận
                            sp = SPs.ThuocXuatkho(objPhieuNhap.IdKhonhap, objPhieuNhapCt.IdThuoc,
                                                          objPhieuNhapCt.NgayHethan, objPhieuNhapCt.GiaNhap, objPhieuNhapCt.GiaBan,
                                                          Utility.DecimaltoDbnull(objPhieuNhapCt.Vat),
                                                          objPhieuNhapCt.SoLuong, objPhieuNhapCt.IdThuockho, objPhieuNhapCt.MaNhacungcap, objPhieuNhapCt.SoLo, objHisDuocProperties.XoaDulieuKhiThuocDaHet ? 1 : 0, errorMessage);
                            sp.Execute();
                            //Cạp nhật lại id_thuockho =-1(giá trị này được update khi xác nhận phiếu). Giá trị id_chuyen cho biết chuyển từ id_thuockho của kho nội trú.
                            new Update(TPhieuNhapxuatthuocChitiet.Schema).Set(TPhieuNhapxuatthuocChitiet.Columns.IdThuockho).EqualTo(-1)
                            .Where(TPhieuNhapxuatthuocChitiet.Columns.IdPhieuchitiet).IsEqualTo(objPhieuNhapCt.IdPhieuchitiet).Execute();
                        }
                        new Update(TPhieuNhapxuatthuoc.Schema)
                            .Set(TPhieuNhapxuatthuoc.Columns.IdNhanvien).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NguoiXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.NgayXacnhan).EqualTo(null)
                            .Set(TPhieuNhapxuatthuoc.Columns.TrangThai).EqualTo(0)
                            .Where(TPhieuNhapxuatthuoc.Columns.IdPhieu).IsEqualTo(objPhieuNhap.IdPhieu)
                            .And(TPhieuNhapxuatthuoc.LoaiPhieuColumn).IsEqualTo(objPhieuNhap.LoaiPhieu).Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                Utility.CatchException("Lỗi khi hủy xác nhận phiếu trả thuốc từ tủ trực về kho lẻ nội trú", ex);
                return ActionResult.Error;
            }
        }
       
        public ActionResult CappnhatPhieucapphatNoitru(TPhieuCapphatNoitru _phieucapphat, List<TPhieuCapphatChitiet> lstPhieuCapphatCt, List<TThuocCapphatChitiet> lstThuocCapphatCt)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {

                        new Update(TPhieuCapphatNoitru.Schema).Set(TPhieuCapphatNoitru.NgayNhapColumn).EqualTo(_phieucapphat.NgayNhap.Date)
                            .Set(TPhieuCapphatNoitru.TuNgayColumn).EqualTo(_phieucapphat.TuNgay.Date)
                            .Set(TPhieuCapphatNoitru.DenNgayColumn).EqualTo(_phieucapphat.DenNgay.Date)
                            .Set(TPhieuCapphatNoitru.IdKhoXuatColumn).EqualTo(_phieucapphat.IdKhoXuat)
                             .Set(TPhieuCapphatNoitru.NgaySuaColumn).EqualTo(_phieucapphat.NgaySua.Value.Date)
                            .Set(TPhieuCapphatNoitru.NguoiSuaColumn).EqualTo(_phieucapphat.NguoiSua)
                            .Where(TPhieuCapphatNoitru.IdCapphatColumn).IsEqualTo(_phieucapphat.IdCapphat).Execute();
                        new Delete().From(TPhieuCapphatChitiet.Schema).Where(TPhieuCapphatChitiet.Columns.IdCapphat).IsEqualTo(
                            _phieucapphat.IdCapphat).Execute();
                        foreach (var _PhieuCapphatCt in lstPhieuCapphatCt)
                        {
                            _PhieuCapphatCt.IdCapphat = _phieucapphat.IdCapphat;
                            _PhieuCapphatCt.IsNew = true;
                            _PhieuCapphatCt.Save();
                        }
                        new Delete().From(TThuocCapphatChitiet.Schema).Where(TThuocCapphatChitiet.Columns.IdCapphat).IsEqualTo(
                           _phieucapphat.IdCapphat).Execute();
                        foreach (var _ThuocCapphatCt in lstThuocCapphatCt)
                        {
                            _ThuocCapphatCt.IdCapphat = _phieucapphat.IdCapphat;
                            _ThuocCapphatCt.IsNew = true;
                            _ThuocCapphatCt.Save();
                        }
                        Scope.Complete();
                    }
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        public ActionResult XoaPhieuCapPhatNoiTru(int ID_CAPPHAT)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        new Delete().From(TPhieuCapphatNoitru.Schema).Where(TPhieuCapphatNoitru.Columns.IdCapphat).IsEqualTo(ID_CAPPHAT).Execute();
                        new Delete().From(TPhieuCapphatChitiet.Schema).Where(TPhieuCapphatChitiet.Columns.IdCapphat).IsEqualTo(ID_CAPPHAT).Execute();
                        new Delete().From(TThuocCapphatChitiet.Schema).Where(TThuocCapphatChitiet.Columns.IdCapphat).IsEqualTo(ID_CAPPHAT).Execute();
                        Scope.Complete();
                    }
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        public ActionResult BenhNhanLinhThuoc13123(int ID_CAPPHAT, int ID_DONTHUOC,long IdChitietdonthuoc, int Trang_thai, int soluongtralai, int thuclinh)
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        new Update(TPhieuCapphatChitiet.Schema)
                            .Set(TPhieuCapphatChitiet.DaLinhColumn.ColumnName).EqualTo(Trang_thai)
                            .Set(TPhieuCapphatChitiet.ThucLinhColumn.ColumnName).EqualTo(thuclinh)
                            .Set(TPhieuCapphatChitiet.SoLuongtralaiColumn.ColumnName).EqualTo(soluongtralai)
                            .Where(TPhieuCapphatChitiet.IdCapphatColumn).IsEqualTo(ID_CAPPHAT)
                            .And(TPhieuCapphatChitiet.IdDonthuocColumn).IsEqualTo(ID_DONTHUOC)
                            .And(TPhieuCapphatChitiet.Columns.IdChitietdonthuoc).IsEqualTo(IdChitietdonthuoc)
                            .Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        public static ActionResult BenhNhanLinhThuoc(int idPhieucapphat,List<int> lstIdDonthuoc, int Trang_thai,ref List<int> lstNoValidData)
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        foreach (int IdDonthuoc in lstIdDonthuoc)
                        {
                            if (Trathuocthua.ThuocNoitruKiemtraThuoctralai(idPhieucapphat, IdDonthuoc))
                            {
                                lstNoValidData.Add(IdDonthuoc);
                                continue;
                            }

                            DataTable dtData = SPs.ThuocNoitruLaychitietdonthuocTheophieulinhthuocnoitru(idPhieucapphat, IdDonthuoc).GetDataSet().Tables[0];
                            foreach (DataRow dr in dtData.Rows)
                            {
                                int ID_CAPPHAT = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdCapphat], -1);
                                int ID_DONTHUOC = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdDonthuoc], -1);
                                int IdChitietdonthuoc = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdChitietdonthuoc], -1);
                                int soluongtralai = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.SoLuongtralai], 0);
                                int SoLuong = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.SoLuong], 0);
                                int thuclinh = SoLuong - soluongtralai;// Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.ThucLinh], -1);
                                if(Trang_thai==0)
                                {
                                    soluongtralai = 0;
                                    thuclinh = 0;
                                }
                                new Update(TPhieuCapphatChitiet.Schema)
                                    .Set(TPhieuCapphatChitiet.DaLinhColumn.ColumnName).EqualTo(Trang_thai)
                                    .Set(TPhieuCapphatChitiet.ThucLinhColumn.ColumnName).EqualTo(thuclinh)
                                    .Set(TPhieuCapphatChitiet.SoLuongtralaiColumn.ColumnName).EqualTo(soluongtralai)
                                    .Set(TPhieuCapphatChitiet.Columns.NguoiSua).EqualTo(globalVariables.UserName)
                                    .Set(TPhieuCapphatChitiet.Columns.NgaySua).EqualTo(globalVariables.SysDate)
                                    .Where(TPhieuCapphatChitiet.IdCapphatColumn).IsEqualTo(ID_CAPPHAT)
                                    .And(TPhieuCapphatChitiet.IdDonthuocColumn).IsEqualTo(ID_DONTHUOC)
                                    .And(TPhieuCapphatChitiet.Columns.IdChitietdonthuoc).IsEqualTo(IdChitietdonthuoc)
                                    .Execute();
                            }
                        }
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        public static ActionResult CapnhatThuclinh(long id_chitiet, int thuc_linh, int so_luongtralai)
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        new Update(TPhieuCapphatChitiet.Schema)
                            .Set(TPhieuCapphatChitiet.Columns.ThucLinh).EqualTo(thuc_linh)
                            .Set(TPhieuCapphatChitiet.Columns.SoLuongtralai).EqualTo(so_luongtralai)
                            .Where(TPhieuCapphatChitiet.Columns.IdChitiet).IsEqualTo(id_chitiet)
                            .Execute();
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
        public static ActionResult BenhNhanLinhThuoc(DataTable dtData,int Trang_thai )
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        foreach (DataRow dr in dtData.Rows)
                        {
                            int ID_CAPPHAT=Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdCapphat],-1);
                            int ID_DONTHUOC=Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdDonthuoc],-1);
                            int SoLuong = Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.SoLuong], 0);
                            int IdChitietdonthuoc=Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.IdChitietdonthuoc],-1);
                            int soluongtralai=Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.SoLuongtralai],0);
                            int thuclinh = SoLuong - soluongtralai;// Utility.Int32Dbnull(dr[TPhieuCapphatChitiet.Columns.ThucLinh], -1);
                            if (Trang_thai == 0)
                            {
                                soluongtralai = 0;
                                thuclinh = 0;
                            }
                            new Update(TPhieuCapphatChitiet.Schema)
                                .Set(TPhieuCapphatChitiet.DaLinhColumn.ColumnName).EqualTo(Trang_thai)
                                .Set(TPhieuCapphatChitiet.ThucLinhColumn.ColumnName).EqualTo(thuclinh)
                                .Set(TPhieuCapphatChitiet.SoLuongtralaiColumn.ColumnName).EqualTo(soluongtralai)
                                .Set(TPhieuCapphatChitiet.Columns.NguoiSua).EqualTo(globalVariables.UserName)
                                .Set(TPhieuCapphatChitiet.Columns.NgaySua).EqualTo(globalVariables.SysDate)
                                .Where(TPhieuCapphatChitiet.IdCapphatColumn).IsEqualTo(ID_CAPPHAT)
                                .And(TPhieuCapphatChitiet.IdDonthuocColumn).IsEqualTo(ID_DONTHUOC)
                                .And(TPhieuCapphatChitiet.Columns.IdChitietdonthuoc).IsEqualTo(IdChitietdonthuoc)
                                .Execute();
                        }
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }
     
        public ActionResult Kiemtratonthuoc(long ID_CAPPHAT, short ID_KHO)
        {
            ActionResult _result = ActionResult.Success;
            try
            {
                TPhieuCapphatChitietCollection lstCapphatchitiet=new Select().From(TPhieuCapphatChitiet.Schema)
                     .Where(TPhieuCapphatChitiet.IdCapphatColumn).IsEqualTo(ID_CAPPHAT).ExecuteAsCollection<TPhieuCapphatChitietCollection>();
                //Xác nhận từng đơn thuốc nội trú
                foreach (TPhieuCapphatChitiet pres in lstCapphatchitiet)
                {
                    if (pres.DaLinh == 0)//Chưa được lĩnh mới kiểm tra
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
        public ActionResult ThemPhieuCapPhatNoiTru(TPhieuCapphatNoitru _phieucapphat,List< TPhieuCapphatChitiet> lstPhieuCapphatCt, List<TThuocCapphatChitiet> lstThuocCapphatCt)
        {
            try
            {
                using (var Scope = new TransactionScope())
                {
                    using (var dbScope = new SharedDbConnectionScope())
                    {
                        _phieucapphat.IsNew = true;
                        _phieucapphat.Save();
                        if (_phieucapphat.IdCapphat > 0)
                        {
                            foreach (var _PhieuCapphatCt in lstPhieuCapphatCt)
                            {
                                _PhieuCapphatCt.IdCapphat = _phieucapphat.IdCapphat;
                                _PhieuCapphatCt.IsNew = true;
                                _PhieuCapphatCt.Save();
                            }
                            foreach (var _ThuocCapphatCt in lstThuocCapphatCt)
                            {
                                _ThuocCapphatCt.IdCapphat = _phieucapphat.IdCapphat;
                                _ThuocCapphatCt.IsNew = true;
                                _ThuocCapphatCt.Save();
                            }
                        }
                       
                    }
                    Scope.Complete();
                    return ActionResult.Success;
                }
            }
            catch (Exception)
            {
                return ActionResult.Error;
            }
        }

    }
}
