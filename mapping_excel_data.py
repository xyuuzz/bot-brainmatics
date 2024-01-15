class MappingExcel:
    df = None

    def __init__(self, df):
        self.df = df

    def mappingData(self):
        namaTraining = self.df.iloc[0, 1]
        tanggalTraining = self.df.iloc[1, 1]
        waktuTraining = self.df.iloc[2, 1]
        LokasiTraining = self.df.iloc[3, 1]
        ruanganTraining = self.df.iloc[4, 1]
        jmlPeserta = self.df.iloc[5, 1]
