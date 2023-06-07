class Meachine:

    

    def __init__(self, maMay,loaiMay, diaChi, soDienThoai):
        self.maMay = maMay;
        self.loaiMay = loaiMay;
        self.diaChi = diaChi;
        self.soDienThoai = soDienThoai;
        self.tienMay = 0;
        self.phiDichVu = 0;

    def upDateMoney(self, themTien):
        self.tienMay += themTien;

    def upDatePhi(self, themTien):
        self.phiDichVu += themTien;
