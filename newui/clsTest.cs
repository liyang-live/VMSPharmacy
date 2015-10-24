 public ActionResult ThanhtoanDonthuoctaiquay(KcbThanhtoan objThanhtoan, KcbDanhsachBenhnhan objBenhnhan,List< KcbThanhtoanChitiet> objArrPaymentDetail, ref int id_thanhtoan, long IdHdonLog, bool Layhoadondo)
        {

            decimal PtramBHYT = 0;
            ///tổng tiền hiện tại truyền vào của lần payment đang thực hiện
            decimal v_dblTongtienDCT = 0;
            ///tổng tiền đã thanh toán
            decimal v_TotalPaymentDetail = 0;
            try
            {
                using (var scope = new TransactionScope())
                {
                    using (var dbscope = new SharedDbConnectionScope())
                    {
                        ///lấy tổng số Payment của mang truyền vào của pay ment hiện tại
                        v_dblTongtienDCT = TongtienKhongTutuc(objArrPaymentDetail);
                        KcbThanhtoanCollection paymentCollection =
                            new KcbThanhtoanController()
                            .FetchByQuery(
                                KcbThanhtoan.CreateQuery()
                                .AddWhere
                                //(KcbThanhtoan.Columns.MaLuotkham, Comparison.Equals, objLuotkham.MaLuotkham).AND
                                (KcbThanhtoan.Columns.IdBenhnhan, Comparison.Equals, objBenhnhan.IdBenhnhan)
                                .AND(KcbThanhtoan.Columns.TrangThai, Comparison.Equals, 0)
                                .AND(KcbThanhtoan.Columns.KieuThanhtoan, Comparison.Equals, 0)
                                .AND(KcbThanhtoan.Columns.TrangThai, Comparison.Equals, 0));
                        //Lấy tổng tiền của các lần thanh toán trước
                        int id_donthuoc = -1;
                        foreach (KcbThanhtoan Payment in paymentCollection)
                        {
                            KcbThanhtoanChitietCollection paymentDetailCollection = new Select().From(KcbThanhtoanChitiet.Schema)
                                .Where(KcbThanhtoanChitiet.Columns.IdThanhtoan).IsEqualTo(Payment.IdThanhtoan)
                                .And(KcbThanhtoanChitiet.Columns.TrangthaiHuy).IsEqualTo(0).ExecuteAsCollection
                                <KcbThanhtoanChitietCollection>();

                            foreach (KcbThanhtoanChitiet paymentDetail in paymentDetailCollection)
                            {
                                if (id_donthuoc == -1) id_donthuoc = paymentDetail.IdPhieu;
                                if (paymentDetail.TuTuc == 0)
                                    v_TotalPaymentDetail += Utility.Int32Dbnull(paymentDetail.SoLuong) *
                                                            Utility.DecimaltoDbnull(paymentDetail.DonGia);

                            }
                        }
                      
                        //LayThongtinPtramBHYT(v_dblTongtienDCT + v_TotalPaymentDetail, objLuotkham, ref PtramBHYT);
                        objThanhtoan.MaThanhtoan = THU_VIEN_CHUNG.TaoMathanhtoan(Convert.ToDateTime(objThanhtoan.NgayThanhtoan));
                        objThanhtoan.IsNew = true;
                        objThanhtoan.Save();
                        if (id_donthuoc == -1) id_donthuoc = objArrPaymentDetail[0].IdPhieu;
                        KcbDonthuoc objDonthuoc = KcbDonthuoc.FetchByID(id_donthuoc);
                        KcbDonthuocChitietCollection lstChitiet = new Select().From(KcbDonthuoc.Schema).Where(KcbDonthuoc.Columns.IdDonthuoc).IsEqualTo(id_donthuoc).ExecuteAsCollection<KcbDonthuocChitietCollection>();
                        ActionResult actionResult = ActionResult.Success;
                        if (objDonthuoc != null && lstChitiet.Count>0)
                        {
                            if (!XuatThuoc.InValiKiemTraDonThuoc(lstChitiet,(byte)0)) return ActionResult.NotEnoughDrugInStock;
                            actionResult = new XuatThuoc().LinhThuocBenhNhan(id_donthuoc, Utility.Int16Dbnull(lstChitiet[0].IdKho, 0), globalVariables.SysDate);
                            switch (actionResult)
                            {
                                case ActionResult.Success:
                                  
                                    break;
                                case ActionResult.Error:
                                    return actionResult;
                            }
                        }
                        //Tính lại Bnhan chi trả và BHYT chi trả
                        //objArrPaymentDetail = THU_VIEN_CHUNG.TinhPhamTramBHYT(objArrPaymentDetail, PtramBHYT);
                        decimal TT_BN = 0m;
                        decimal TT_BHYT = 0m;
                        decimal TT_Chietkhau_Chitiet = 0m;
                        foreach (KcbThanhtoanChitiet objThanhtoanDetail in objArrPaymentDetail)
                        {
                            TT_BN += (objThanhtoanDetail.BnhanChitra + objThanhtoanDetail.PhuThu) * objThanhtoanDetail.SoLuong;
                            TT_BHYT += objThanhtoanDetail.BhytChitra * objThanhtoanDetail.SoLuong;
                            TT_Chietkhau_Chitiet += Utility.DecimaltoDbnull(objThanhtoanDetail.TienChietkhau, 0);
                            objThanhtoanDetail.IdThanhtoan = Utility.Int32Dbnull(objThanhtoan.IdThanhtoan, -1);
                            objThanhtoanDetail.IsNew = true;
                            objThanhtoanDetail.Save();
                            UpdatePaymentStatus(objThanhtoan, objThanhtoanDetail);
                        }

                        #region Hoadondo

                        if (Layhoadondo)
                        {
                            int record = -1;
                            if (IdHdonLog > 0)
                            {
                                record =
                                    new Delete().From(HoadonLog.Schema)
                                        .Where(HoadonLog.Columns.IdHdonLog)
                                        .IsEqualTo(IdHdonLog)
                                        .Execute();
                                if (record <= 0)
                                {
                                    Utility.ShowMsg("Có lỗi trong quá trình xóa thông tin serie hóa đơn đã hủy để cấp lại cho lần thanh toán này.");
                                    return ActionResult.Error;
                                }
                            }
                            var obj = new HoadonLog();
                            obj.IdThanhtoan = objThanhtoan.IdThanhtoan;
                            obj.TongTien = objThanhtoan.TongTien - Utility.DecimaltoDbnull(objThanhtoan.TongtienChietkhau, 0);
                            obj.IdBenhnhan = objThanhtoan.IdBenhnhan;
                            obj.MaLuotkham = objThanhtoan.MaLuotkham;
                            obj.MauHoadon = objThanhtoan.MauHoadon;
                            obj.KiHieu = objThanhtoan.KiHieu;
                            obj.IdCapphat = objThanhtoan.IdCapphat.Value;
                            obj.MaQuyen = objThanhtoan.MaQuyen;
                            obj.Serie = objThanhtoan.Serie;
                            obj.MaNhanvien = globalVariables.UserName;
                            obj.MaLydo = "0";
                            obj.NgayIn = globalVariables.SysDate;
                            obj.TrangThai = 0;
                            obj.IsNew = true;
                            obj.Save();
                            IdHdonLog = obj.IdHdonLog;//Để update lại vào bảng thanh toán
                            new Update(HoadonCapphat.Schema).Set(HoadonCapphat.Columns.SerieHientai)
                                .EqualTo(objThanhtoan.Serie)
                                .Set(HoadonCapphat.Columns.TrangThai).EqualTo(1)
                                .Where(HoadonCapphat.Columns.IdCapphat).IsEqualTo(obj.IdCapphat)
                                .Execute();
                        }
                        #endregion
                        KcbPhieuthu objPhieuthu = new KcbPhieuthu();
                        objPhieuthu.IdBenhnhan = objThanhtoan.IdBenhnhan;
                        objPhieuthu.MaLuotkham = objThanhtoan.MaLuotkham;
                        objPhieuthu.IdThanhtoan = objThanhtoan.IdThanhtoan;
                        objPhieuthu.MaPhieuthu = THU_VIEN_CHUNG.GetMaPhieuThu(globalVariables.SysDate, 0);
                        objPhieuthu.SoluongChungtugoc = 1;
                        objPhieuthu.LoaiPhieuthu = Convert.ToByte(0);
                        objPhieuthu.NgayThuchien = globalVariables.SysDate;
                        objPhieuthu.SoTien = TT_BN - TT_Chietkhau_Chitiet;
                        objPhieuthu.SotienGoc = TT_BN;
                        objPhieuthu.MaLydoChietkhau = objThanhtoan.MaLydoChietkhau;
                        objPhieuthu.TienChietkhauchitiet = TT_Chietkhau_Chitiet;
                        objPhieuthu.TienChietkhau = objThanhtoan.TongtienChietkhau;
                        objPhieuthu.TienChietkhauhoadon = objPhieuthu.TienChietkhau - objPhieuthu.TienChietkhauchitiet;
                        objPhieuthu.NguoiNop = globalVariables.UserName;
                        objPhieuthu.TaikhoanCo = "";
                        objPhieuthu.TaikhoanNo = "";
                        objPhieuthu.LydoNop = "Thu tiền bệnh nhân";
                        objPhieuthu.IdKhoaThuchien = globalVariables.idKhoatheoMay;
                        objPhieuthu.IdNhanvien = globalVariables.gv_intIDNhanvien;
                        objPhieuthu.IsNew = true;
                        objPhieuthu.Save();

                        new Update(KcbThanhtoan.Schema)
                        .Set(KcbThanhtoan.Columns.TongTien).EqualTo(TT_BHYT + TT_BN)
                        .Set(KcbThanhtoan.Columns.BnhanChitra).EqualTo(TT_BN)
                        .Set(KcbThanhtoan.Columns.BhytChitra).EqualTo(TT_BHYT)
                        .Set(KcbThanhtoan.Columns.IdHdonLog).EqualTo(IdHdonLog)
                        .Where(KcbThanhtoan.Columns.IdThanhtoan).IsEqualTo(objThanhtoan.IdThanhtoan).Execute();
                    }
                    scope.Complete();
                    id_thanhtoan = Utility.Int32Dbnull(objThanhtoan.IdThanhtoan, -1);
                    return ActionResult.Success;
                }
            }
            catch (Exception ex)
            {
                log.Error("Loi thuc hien thanh toan:" + ex.ToString());
                return ActionResult.Error;
            }

        }