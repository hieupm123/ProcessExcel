from meachine import Meachine


class Company:

    

    def __init__(self, maCty, tenCongTy):
        self.maCty = maCty;
        self.tenCongTy = tenCongTy;
        self.phiCongTy = 0;
        self.tienCongTy = 0;
        self.meachines = [];

    def upDateMeachine(self, meachine):
        self.meachines.append(meachine);

    def upDateTienVaPhiCty(self):
        self.phiCongTy = 0;
        self.tienCongTy = 0;
        for i in self.meachines:
            self.tienCongTy += i.tienMay;
            self.phiCongTy += i.phiDichVu;
