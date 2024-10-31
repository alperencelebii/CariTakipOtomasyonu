using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CariTakip
{
    public partial class XtraAnaSayfa : DevExpress.XtraEditors.XtraForm
    {
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\Data\db.mdb");
        public XtraAnaSayfa af;
        public XtraCari ca;
        public XtraCariHareket ch;
        public XtraStok st;
        private Point barManager1;

        private PrintDocument printDocument = new PrintDocument();
        private PrintDocument genelPrintDocument = new PrintDocument();
        private int currentRow = 0;

        public XtraAnaSayfa()
        {
            InitializeComponent();
            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
            genelPrintDocument.PrintPage += new PrintPageEventHandler(genelPrintDocument_PrintPage);
        }

        private void XtraAnaSayfa_Load(object sender, EventArgs e)
        {
            göster();
            InitializeButtons(); // Yeni butonları burada tanımlıyoruz
        }
        private void ChangeDatabaseButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // Kullanıcıya dosya seçtirmek için OpenFileDialog oluşturuyoruz
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "MDB Files|*.mdb",
                    Title = "Veritabanı Dosyasını Seç",
                    FileName = "db.mdb"
                };

                // Eğer kullanıcı bir dosya seçtiyse
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Yeni veritabanı dosyası yolunu alıyoruz
                    string selectedDatabasePath = openFileDialog.FileName;

                    // Bağlantı açık mı kontrol edelim
                    if (baglanti != null && baglanti.State == ConnectionState.Open)
                    {
                        // Eğer bağlantı zaten açıksa, önce bağlantıyı kapatalım ve kaynakları serbest bırakalım
                        baglanti.Close();
                        baglanti.Dispose();
                        baglanti = null; // Eski bağlantıyı tamamen serbest bırakıyoruz
                        GC.Collect();  // Garbage Collector'u çağırıyoruz
                        GC.WaitForPendingFinalizers(); // Kaynakları tamamen serbest bırakmak için finalize çağrısını bekliyoruz
                    }

                    // Yeni bağlantı dizini oluşturuyoruz
                    baglanti = new OleDbConnection($@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={selectedDatabasePath}");

                    // Bağlantıyı açıyoruz ve işlemleri yeni veritabanıyla sürdüreceğiz
                    baglanti.Open();

                    // Kullanıcıya başarı mesajı gösteriyoruz
                    MessageBox.Show("Veritabanı başarıyla değiştirildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Yeni veritabanıyla işlemleri devam ettir
                    göster();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veritabanı değiştirilirken bir hata oluştu: " + ex.Message);
            }
        }
        public void göster()
        {
            // Bağlantı açık mı kontrol edelim
            if (baglanti.State == ConnectionState.Open)
            {
                baglanti.Close(); // Bağlantı açıksa önce kapat
            }

            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARI where DURUM=true order by ADISOYADI", baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);

            // DataTable'da BORC ve VERILEN sütunlarını ekleyelim
            if (!da.Columns.Contains("BORC"))
            {
                da.Columns.Add("BORC", typeof(double));
            }

            if (!da.Columns.Contains("VERILEN"))
            {
                da.Columns.Add("VERILEN", typeof(double));
            }

            // Veritabanı işlemlerini tek bir bağlantı ile yapıyoruz
            baglanti.Open();
            int rowIndex = 1; // Sıra numarası için değişkeni tanımlıyoruz
            foreach (DataRow row in da.Rows)
            {
                int cariId = Convert.ToInt32(row["ID"]);

                // Cari borçları ve alacakları topluca hesapla
                OleDbCommand command = new OleDbCommand(
                    "SELECT SUM(BORC) AS TotalBorc, SUM(ALACAK) AS TotalVerilen FROM TBLCARIHAREKET WHERE CARI_ID = @CARI_ID", baglanti);
                command.Parameters.AddWithValue("@CARI_ID", cariId);

                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        double toplamBorc = reader["TotalBorc"] != DBNull.Value ? Convert.ToDouble(reader["TotalBorc"]) : 0;
                        double toplamVerilen = reader["TotalVerilen"] != DBNull.Value ? Convert.ToDouble(reader["TotalVerilen"]) : 0;

                        // Hesaplanan değerleri tabloya ekle
                        row["BORC"] = toplamBorc;
                        row["VERILEN"] = toplamVerilen;
                    }
                }

                // İsimlerin başına numara ekleyelim
                row["ADISOYADI"] = $"{rowIndex}. {row["ADISOYADI"]}";
                rowIndex++; // Sıra numarasını artırıyoruz
            }
            baglanti.Close();

            gC1.DataSource = da;

            // Borç ve Verilen sütunları kontrol edip ekle (eklenmemişse)
            if (gridView1.Columns["BORC"] == null)
            {
                DevExpress.XtraGrid.Columns.GridColumn columnBorc = new DevExpress.XtraGrid.Columns.GridColumn();
                columnBorc.FieldName = "BORC";
                columnBorc.Caption = "Borç";
                columnBorc.Visible = true;
                columnBorc.VisibleIndex = gridView1.Columns["ADISOYADI"].VisibleIndex + 1;
                columnBorc.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                columnBorc.DisplayFormat.FormatString = "n2";
                gridView1.Columns.Add(columnBorc);
            }

            if (gridView1.Columns["VERILEN"] == null)
            {
                DevExpress.XtraGrid.Columns.GridColumn columnVerilen = new DevExpress.XtraGrid.Columns.GridColumn();
                columnVerilen.FieldName = "VERILEN";
                columnVerilen.Caption = "Verilen";
                columnVerilen.Visible = true;
                columnVerilen.VisibleIndex = gridView1.Columns["BORC"].VisibleIndex + 1;
                columnVerilen.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                columnVerilen.DisplayFormat.FormatString = "n2";
                gridView1.Columns.Add(columnVerilen);
            }

            // Bakiye sütununu en sağa taşı
            gridView1.Columns["BAKIYE"].VisibleIndex = gridView1.Columns.Count - 1;

            // Bakiye ve Adı Soyadı sütunlarının genişliklerini ayarla
            gridView1.Columns["BAKIYE"].Width = 300;
            gridView1.Columns["ADISOYADI"].Width = 300;

            // Sola hizalama ayarı
            foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView1.Columns)
            {
                column.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            }
        }

        private void InitializeButtons()
        {
            DevExpress.XtraBars.BarButtonItem changeDatabaseButton = new DevExpress.XtraBars.BarButtonItem();
            changeDatabaseButton.Caption = "Veritabanı Değiştir";
            changeDatabaseButton.ItemClick += ChangeDatabaseButton_ItemClick;

            // Veritabanı değiştir butonu için resim ekleme (isteğe bağlı)
            string imagePathDatabase = System.IO.Path.Combine(Application.StartupPath, "images\\database.png");
            if (System.IO.File.Exists(imagePathDatabase))
            {
                changeDatabaseButton.LargeGlyph = Image.FromFile(imagePathDatabase);
            }

            // Mevcut bir gruba butonları ekleyin
            ribbonPageGroup1.ItemLinks.Add(changeDatabaseButton);
            // Yazdırma butonunu oluşturma
            DevExpress.XtraBars.BarButtonItem printButton = new DevExpress.XtraBars.BarButtonItem();
            printButton.Caption = "Genel Yazdır";
            printButton.ItemClick += PrintButton_ItemClick;

            // Genel Yazdır butonunu oluşturma
            DevExpress.XtraBars.BarButtonItem genelPrintButton = new DevExpress.XtraBars.BarButtonItem();
            genelPrintButton.Caption = "Yazdır";
            genelPrintButton.ItemClick += GenelPrintButton_ItemClick;

            // Excel'e Aktar butonunu oluşturma
            DevExpress.XtraBars.BarButtonItem exportToExcelButton = new DevExpress.XtraBars.BarButtonItem();
            exportToExcelButton.Caption = "Genel Excel'e Aktar";
            exportToExcelButton.ItemClick += ExportToExcelButton_ItemClick;

            DevExpress.XtraBars.BarButtonItem exportFilteredExcelButton = new DevExpress.XtraBars.BarButtonItem();
            exportFilteredExcelButton.Caption = "Excel'e Aktar";
            exportFilteredExcelButton.ItemClick += ExportFilteredExcelButton_ItemClick;


            // Yazdır butonları için resim yolu
            string imagePathPrint = System.IO.Path.Combine(Application.StartupPath, "images\\yazdir.png");

            // Excel butonu için resim yolu
            string imagePathExcel = System.IO.Path.Combine(Application.StartupPath, "images\\excel.png");
            string newImagePath = System.IO.Path.Combine(Application.StartupPath, "images\\excel.png");
            string imagepathveri = @"images\\veri.png"; // Resim yolu

            if (System.IO.File.Exists(imagepathveri))
            {
                changeDatabaseButton.LargeGlyph = Image.FromFile(imagepathveri); // Resmi butona ekle
            }
            else
            {
                MessageBox.Show("Veritabanı resmi bulunamadı: " + imagePathDatabase);
            }
            if (System.IO.File.Exists(imagePathPrint))
            {
                // Yazdırma butonlarına büyük bir resim ekleme
                printButton.LargeGlyph = Image.FromFile(imagePathPrint);
                genelPrintButton.LargeGlyph = Image.FromFile(imagePathPrint);
            }
            else
            {
                MessageBox.Show("Yazdırma resim dosyası bulunamadı: " + imagePathPrint);
            }

            if (System.IO.File.Exists(imagePathExcel))
            {
                // Excel butonuna farklı bir resim ekleme
                exportToExcelButton.LargeGlyph = Image.FromFile(imagePathExcel);
            }
            else
            {
                MessageBox.Show("Excel resim dosyası bulunamadı: " + imagePathExcel);
            }
            if (System.IO.File.Exists(newImagePath))
            {
                // Excel butonuna farklı bir resim ekleme
                exportFilteredExcelButton.LargeGlyph = Image.FromFile(newImagePath);
            }

            // Mevcut bir gruba butonları ekleyin
            ribbonPageGroup1.ItemLinks.Add(printButton);
            ribbonPageGroup1.ItemLinks.Add(genelPrintButton);
            ribbonPageGroup1.ItemLinks.Add(exportToExcelButton);
            ribbonPageGroup1.ItemLinks.Add(exportFilteredExcelButton); // Excel butonunu gruba ekliyoruz
        }
        private void ExportFilteredExcelButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // SaveFileDialog ile dosya kaydedilecek yeri belirleyin
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Excel Dosyası Kaydet",
                    FileName = "FilteredCariHesapEkstre_Bakiye.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // GridControl'deki verileri DataTable'a çekiyoruz
                    DataTable dt = gC1.DataSource as DataTable; // Veri kaynağını DataTable olarak alıyoruz

                    if (dt == null) // Eğer DataTable boşsa bir uyarı ver
                    {
                        MessageBox.Show("Veri bulunamadı. Lütfen geçerli bir tablo yükleyin.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return; // İşlemi burada durdur
                    }

                    // BAKIYE sütunu 0 olmayanları filtreleyelim
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        // Eğer BAKIYE sütunu varsa ve değer 0 değilse
                        if (dt.Columns.Contains("BAKIYE") && dt.Rows[i]["BAKIYE"] != DBNull.Value)
                        {
                            double bakiyeDegeri = Convert.ToDouble(dt.Rows[i]["BAKIYE"]);
                            if (bakiyeDegeri == 0)
                            {
                                dt.Rows.RemoveAt(i); // Bakiye 0 olanları çıkarıyoruz
                            }
                        }
                        else
                        {
                            // Eğer BAKIYE sütunu yoksa veya değer null ise satırı çıkar
                            dt.Rows.RemoveAt(i);
                        }
                    }

                    // Geçici GridControl'e veri kaynağını set edip dışa aktaralım
                    GridControl tempGridControl = new GridControl();
                    tempGridControl.DataSource = dt;
                    tempGridControl.ExportToXlsx(saveFileDialog.FileName);

                    MessageBox.Show("Bakiye 0 olmayan veriler başarıyla Excel'e aktarıldı.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel'e aktarma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void PrintButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // İlk yazdır butonu tüm verileri alacak
            toplamBorc = 0;
            toplamVerilen = 0;
            toplamBakiye = 0;
            currentRow = 0;
            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = printDocument;
            printPreviewDialog.ShowDialog();
        }

        private void GenelPrintButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // Genel yazdır butonu sadece bakiye 0 olmayanları alacak
            currentRow = 0;
            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = genelPrintDocument;
            printPreviewDialog.ShowDialog();
        }

        // Toplamları biriktirmek için sınıf seviyesinde tanımlamalar
        private double toplamBorc = 0;
        private double toplamVerilen = 0;
        private double toplamBakiye = 0;

        void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                // Yazdırma işlemi başladığında toplamları sıfırla
                if (currentRow == 0) // İlk sayfa kontrolü
                {
                    toplamBorc = 0;
                    toplamVerilen = 0;
                    toplamBakiye = 0;
                }

                Font font = new Font("Arial", 7); // Yazı tipi
                Font titleFont = new Font("Arial", 10, FontStyle.Bold); // Başlık yazı tipi
                Font headerFont = new Font("Arial", 12, FontStyle.Bold); // Başlık yazı tipi
                Font groupFont = new Font("Arial", 9, FontStyle.Bold); // Grup yazı tipi

                float lineHeight = font.GetHeight() + 1; // Satır yüksekliği

                // Sayfa genişliği ve kenar boşlukları
                float pageWidth = e.MarginBounds.Width;
                float pageLeft = e.MarginBounds.Left;
                float pageRight = e.MarginBounds.Right;

                // Sütun sayısı
                int columnCount = 4;

                // Sütun genişliğini hesaplayalım
                float columnWidth = pageWidth / columnCount;

                // Sütunların X koordinatlarını belirleyelim
                float column1X = pageLeft;
                float column2X = column1X + columnWidth;
                float column3X = column2X + columnWidth;
                float column4X = column3X + columnWidth;

                float y = 30;

                // Sayfa başlığı
                string header = "Genel Cari Hesap Ekstre Raporu";
                float headerWidth = e.Graphics.MeasureString(header, headerFont).Width;
                e.Graphics.DrawString(header, headerFont, Brushes.Black, (pageRight + pageLeft - headerWidth) / 2, y);
                y += headerFont.GetHeight() + 5;

                // Alt çizgi
                float underlinePadding = 10;
                float underlineStartX = (pageRight + pageLeft - headerWidth) / 2 + underlinePadding;
                float underlineEndX = (pageRight + pageLeft + headerWidth) / 2 - underlinePadding;

                e.Graphics.DrawLine(Pens.Black, underlineStartX, y, underlineEndX, y);
                y += 10;

                // Mavi arka plan için genişletilmiş dikdörtgen
                RectangleF headerRect = new RectangleF(pageLeft, y, pageWidth, 35);
                e.Graphics.FillRectangle(Brushes.SteelBlue, headerRect);

                // Cari ismini ortala
                string cariIsmi = "Gökçe Tarım";
                float cariIsmiWidth = e.Graphics.MeasureString(cariIsmi, groupFont).Width;
                e.Graphics.DrawString(cariIsmi, groupFont, Brushes.White, (pageRight + pageLeft - cariIsmiWidth) / 2, y + 5);

                y += 20;

                // Tablo başlıkları hizalı
                StringFormat headerFormat = new StringFormat();
                headerFormat.Alignment = StringAlignment.Center;

                // Başlıkların koordinatlarını ayarlayalım
                float adiSoyadiX = 50f;
                float borcX = 290f;
                float verilenX = 460f;
                float bakiyeX = 600f;

                // Başlıkları çizelim
                e.Graphics.DrawString("Adı Soyadı", font, Brushes.White, new RectangleF(adiSoyadiX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Borç", font, Brushes.White, new RectangleF(borcX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Verilen", font, Brushes.White, new RectangleF(verilenX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Bakiye", font, Brushes.White, new RectangleF(bakiyeX, y, columnWidth, lineHeight), headerFormat);

                y += headerRect.Height - 20;

                // Veriler burada hizalanacak
                DataTable dataTable = (DataTable)gC1.DataSource;

                // Metinler için sola hizalama
                StringFormat leftAlignFormat = new StringFormat();
                leftAlignFormat.Alignment = StringAlignment.Near;

                // Sayısal veriler için sağa hizalama
                StringFormat rightAlignFormat = new StringFormat();
                rightAlignFormat.Alignment = StringAlignment.Far;

                double sayfaToplamBorc = 0;
                double sayfaToplamVerilen = 0;
                double sayfaToplamBakiye = 0;

                while (currentRow < dataTable.Rows.Count)
                {
                    DataRow row = dataTable.Rows[currentRow];

                    if (y + lineHeight > e.MarginBounds.Bottom - 100) // Sayfanın sonuna yaklaşıyorsak yeni sayfaya geçelim
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    string adiSoyadi = row["ADISOYADI"].ToString();
                    double borc = Convert.ToDouble(row["BORC"]);
                    double verilen = Convert.ToDouble(row["VERILEN"]);
                    double bakiye = Convert.ToDouble(row["BAKIYE"]);

                    // Sayfa toplamlarını hesapla
                    sayfaToplamBorc += borc;
                    sayfaToplamVerilen += verilen;
                    sayfaToplamBakiye += bakiye;

                    // Global toplamları biriktir
                    toplamBorc += borc;
                    toplamVerilen += verilen;
                    toplamBakiye += bakiye;

                    // Sütunları hizalı olarak çizelim
                    e.Graphics.DrawString(adiSoyadi, font, Brushes.Black, new RectangleF(column1X, y, columnWidth, lineHeight), leftAlignFormat);
                    e.Graphics.DrawString(borc.ToString("N2"), font, Brushes.Black, new RectangleF(column2X, y, columnWidth, lineHeight), rightAlignFormat);
                    e.Graphics.DrawString(verilen.ToString("N2"), font, Brushes.Black, new RectangleF(column3X, y, columnWidth, lineHeight), rightAlignFormat);
                    e.Graphics.DrawString(bakiye.ToString("N2"), font, Brushes.Black, new RectangleF(column4X, y, columnWidth, lineHeight), rightAlignFormat);

                    // Satırın altına noktalı çizgi ekleyelim
                    Pen dottedPen = new Pen(Color.Black);
                    dottedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                    e.Graphics.DrawLine(dottedPen, column1X, y + lineHeight, pageRight, y + lineHeight);

                    y += lineHeight + 4; // Satır aralığını biraz artırıyoruz
                    currentRow++;
                }

                // Eğer son sayfadaysak, toplamları gösterelim
                if (!e.HasMorePages)
                {
                    // Cari bilgilerinin altına çizgi çizelim
                    y += lineHeight;
                    e.Graphics.DrawLine(Pens.Black, e.MarginBounds.Left, y, e.MarginBounds.Right, y);
                    y += 20;

                    float boxWidth = 200;
                    float boxHeight = 30;
                    float spacing = 10;
                    float boxY = y + 20;
                    float boxX = e.MarginBounds.Right - (boxWidth * 3 + spacing * 2);

                    // Borç kutusu
                    e.Graphics.DrawString("BORÇ", font, Brushes.Black, boxX, boxY - 20);
                    RectangleF borcRect = new RectangleF(boxX, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, borcRect);
                    e.Graphics.DrawRectangle(Pens.Black, borcRect.X, borcRect.Y, borcRect.Width, borcRect.Height);
                    e.Graphics.DrawString(toplamBorc.ToString("N2"), titleFont, Brushes.White, borcRect.X + 10, borcRect.Y + 5);

                    // Verilen kutusu
                    e.Graphics.DrawString("VERİLEN", font, Brushes.Black, borcRect.Right + spacing, boxY - 20);
                    RectangleF alacakRect = new RectangleF(borcRect.Right + spacing, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, alacakRect);
                    e.Graphics.DrawRectangle(Pens.Black, alacakRect.X, alacakRect.Y, alacakRect.Width, alacakRect.Height);
                    e.Graphics.DrawString(toplamVerilen.ToString("N2"), titleFont, Brushes.White, alacakRect.X + 10, alacakRect.Y + 5);

                    // Bakiye kutusu
                    e.Graphics.DrawString("BAKİYE", font, Brushes.Black, alacakRect.Right + spacing, boxY - 20);
                    RectangleF toplamRect = new RectangleF(alacakRect.Right + spacing, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, toplamRect);
                    e.Graphics.DrawRectangle(Pens.Black, toplamRect.X, toplamRect.Y, toplamRect.Width, toplamRect.Height);
                    e.Graphics.DrawString(toplamBakiye.ToString("N2"), titleFont, Brushes.White, toplamRect.X + 10, toplamRect.Y + 5);

                    e.HasMorePages = false; // Son sayfa
                    currentRow = 0; // Tüm sayfalar yazdırıldıktan sonra sıfırla
                }
                else
                {
                    // Eğer son sayfa değilse, toplamları sıfırlamayın ve sayfaya devam edin
                    e.HasMorePages = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yazdırma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }

        double globalToplamBorc = 0;
        double globalToplamVerilen = 0;
        double globalToplamBakiye = 0;

        void genelPrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                // Toplamları her yazdırma işlemine başlamadan önce sıfırla
                if (currentRow == 0)
                {
                    globalToplamBorc = 0;
                    globalToplamVerilen = 0;
                    globalToplamBakiye = 0;
                }

                Font font = new Font("Arial", 7);
                Font titleFont = new Font("Arial", 10, FontStyle.Bold);
                Font headerFont = new Font("Arial", 12, FontStyle.Bold);
                Font groupFont = new Font("Arial", 9, FontStyle.Bold);

                float lineHeight = font.GetHeight() + 1;

                // Sayfa genişliği ve kenar boşlukları
                float pageWidth = e.MarginBounds.Width;
                float pageLeft = e.MarginBounds.Left;
                float pageRight = e.MarginBounds.Right;

                // Sütun sayısı
                int columnCount = 4;

                // Sütun genişliğini sayfa genişliğine göre hesaplayalım
                float columnWidth = pageWidth / columnCount;

                // Sütunların X koordinatlarını ayarlayalım
                float column1X = pageLeft;
                float column2X = column1X + columnWidth;
                float column3X = column2X + columnWidth;
                float column4X = column3X + columnWidth;

                float y = 30;

                // Sayfa başlığı
                string header = "Cari Hesap Ekstre Raporu";
                float headerWidth = e.Graphics.MeasureString(header, headerFont).Width;
                e.Graphics.DrawString(header, headerFont, Brushes.Black, (pageRight + pageLeft - headerWidth) / 2, y);
                y += headerFont.GetHeight() + 5;

                // Alt çizgi
                float underlinePadding = 10;
                float underlineStartX = (pageRight + pageLeft - headerWidth) / 2 + underlinePadding;
                float underlineEndX = (pageRight + pageLeft + headerWidth) / 2 - underlinePadding;

                e.Graphics.DrawLine(Pens.Black, underlineStartX, y, underlineEndX, y);
                y += 10;

                // Mavi arka plan için genişletilmiş dikdörtgen
                RectangleF headerRect = new RectangleF(pageLeft, y, pageWidth, 35);
                e.Graphics.FillRectangle(Brushes.SteelBlue, headerRect);

                // Cari ismini ortala
                string cariIsmi = "Gökçe Tarım";
                float cariIsmiWidth = e.Graphics.MeasureString(cariIsmi, groupFont).Width;
                e.Graphics.DrawString(cariIsmi, groupFont, Brushes.White, (pageRight + pageLeft - cariIsmiWidth) / 2, y + 5);

                y += 20;

                // Tablo başlıkları hizalı
                StringFormat headerFormat = new StringFormat();
                headerFormat.Alignment = StringAlignment.Center;

                // Her bir başlık için manuel olarak X koordinatını ayarlayalım
                float adiSoyadiX = 50f;
                float borcX = 290f;
                float verilenX = 460f;
                float bakiyeX = 600f;

                // Başlıkları çizelim
                e.Graphics.DrawString("Adı Soyadı", font, Brushes.White, new RectangleF(adiSoyadiX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Borç", font, Brushes.White, new RectangleF(borcX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Verilen", font, Brushes.White, new RectangleF(verilenX, y, columnWidth, lineHeight), headerFormat);
                e.Graphics.DrawString("Bakiye", font, Brushes.White, new RectangleF(bakiyeX, y, columnWidth, lineHeight), headerFormat);

                y += headerRect.Height - 20;

                // Veriler burada hizalanacak
                DataTable dataTable = (DataTable)gC1.DataSource;

                // Metinler için sola hizalama
                StringFormat leftAlignFormat = new StringFormat();
                leftAlignFormat.Alignment = StringAlignment.Near;

                // Sayısal veriler için sağa hizalama
                StringFormat rightAlignFormat = new StringFormat();
                rightAlignFormat.Alignment = StringAlignment.Far;

                double sayfaToplamBorc = 0;
                double sayfaToplamVerilen = 0;
                double sayfaToplamBakiye = 0;

                while (currentRow < dataTable.Rows.Count)
                {
                    DataRow row = dataTable.Rows[currentRow];

                    double bakiye = Convert.ToDouble(row["BAKIYE"]);
                    if (bakiye == 0)
                    {
                        currentRow++;
                        continue; // Bakiye 0 ise bu satırı atla
                    }

                    if (y + lineHeight > e.MarginBounds.Bottom - 100)
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    string adiSoyadi = row["ADISOYADI"].ToString();
                    double borc = Convert.ToDouble(row["BORC"]);
                    double verilen = Convert.ToDouble(row["VERILEN"]);

                    // Sayfa toplamlarını biriktiriyoruz
                    sayfaToplamBorc += borc;
                    sayfaToplamVerilen += verilen;
                    sayfaToplamBakiye += bakiye;

                    // Global toplamları biriktiriyoruz (sadece bir kez topluyoruz)
                    globalToplamBorc += borc;
                    globalToplamVerilen += verilen;
                    globalToplamBakiye += bakiye;

                    // Sütunları hizalı çizelim
                    e.Graphics.DrawString(adiSoyadi, font, Brushes.Black, new RectangleF(column1X, y, columnWidth, lineHeight), leftAlignFormat);
                    e.Graphics.DrawString(borc.ToString("N2"), font, Brushes.Black, new RectangleF(column2X, y, columnWidth, lineHeight), rightAlignFormat);
                    e.Graphics.DrawString(verilen.ToString("N2"), font, Brushes.Black, new RectangleF(column3X, y, columnWidth, lineHeight), rightAlignFormat);
                    e.Graphics.DrawString(bakiye.ToString("N2"), font, Brushes.Black, new RectangleF(column4X, y, columnWidth, lineHeight), rightAlignFormat);

                    // Satırın altına noktalı bir çizgi ekliyoruz
                    Pen dottedPen = new Pen(Color.Black);
                    dottedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                    e.Graphics.DrawLine(dottedPen, column1X, y + lineHeight, pageRight, y + lineHeight);

                    y += lineHeight + 4;
                    currentRow++;
                }

                // Eğer son sayfadaysak, tüm sayfa toplamlarını gösteren dikdörtgenler ekle
                if (!e.HasMorePages)
                {
                    // Cari bilgilerinin altına bir çizgi çiziyoruz
                    y += lineHeight;
                    e.Graphics.DrawLine(Pens.Black, e.MarginBounds.Left, y, e.MarginBounds.Right, y);
                    y += 20;

                    float boxWidth = 200;
                    float boxHeight = 30;
                    float spacing = 10;
                    float boxY = y + 20;
                    float boxX = e.MarginBounds.Right - (boxWidth * 3 + spacing * 2);

                    // Borç başlığı ve kutusu
                    e.Graphics.DrawString("BORÇ", font, Brushes.Black, boxX, boxY - 20);
                    RectangleF borcRect = new RectangleF(boxX, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, borcRect);
                    e.Graphics.DrawRectangle(Pens.Black, borcRect.X, borcRect.Y, borcRect.Width, borcRect.Height);
                    e.Graphics.DrawString(globalToplamBorc.ToString("N2"), titleFont, Brushes.White, borcRect.X + 10, borcRect.Y + 5);

                    // Verilen başlığı ve kutusu
                    e.Graphics.DrawString("VERİLEN", font, Brushes.Black, borcRect.Right + spacing, boxY - 20);
                    RectangleF alacakRect = new RectangleF(borcRect.Right + spacing, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, alacakRect);
                    e.Graphics.DrawRectangle(Pens.Black, alacakRect.X, alacakRect.Y, alacakRect.Width, alacakRect.Height);
                    e.Graphics.DrawString(globalToplamVerilen.ToString("N2"), titleFont, Brushes.White, alacakRect.X + 10, alacakRect.Y + 5);

                    // Bakiye başlığı ve kutusu
                    e.Graphics.DrawString("BAKİYE", font, Brushes.Black, alacakRect.Right + spacing, boxY - 20);
                    RectangleF toplamRect = new RectangleF(alacakRect.Right + spacing, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, toplamRect);
                    e.Graphics.DrawRectangle(Pens.Black, toplamRect.X, toplamRect.Y, toplamRect.Width, toplamRect.Height);
                    e.Graphics.DrawString(globalToplamBakiye.ToString("N2"), titleFont, Brushes.White, toplamRect.X + 10, toplamRect.Y + 5);

                    e.HasMorePages = false;
                    currentRow = 0; // Tüm sayfalar yazdırıldıktan sonra sıfırla
                }
                else
                {
                    e.HasMorePages = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yazdırma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void BtnCariEkle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // Güncel veritabanına göre işlem yap
                ca = new XtraCari();
                ca.af = this;
                ca.baglanti = this.baglanti; // Veritabanı bağlantısını ilgili form'a geçiriyoruz
                ca.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cari eklenirken bir hata oluştu: " + ex.Message);
            }
        }

        private void BtnKapat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }

        private void BtnStokEkle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // Güncel veritabanına göre stok ekleme işlemi yap
                st = new XtraStok();
                st.baglanti = this.baglanti; // Bağlantıyı buraya geçiriyoruz
                st.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Stok eklenirken bir hata oluştu: " + ex.Message);
            }
        }

        private void BtnİşlemYap_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // Seçilen satırdaki ID'yi al ve yeni formda güncel bağlantıyı kullan
                ch = new XtraCariHareket();
                ch.af = this;
                ch.baglanti = this.baglanti; // Bağlantıyı buraya geçiriyoruz
                ch.id = Convert.ToInt32(gridView1.GetFocusedRowCellValue("ID").ToString());
                ch.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("İşlem yapılırken bir hata oluştu: " + ex.Message);
            }
        }
        private void ExportToExcelButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // SaveFileDialog ile dosya kaydedilecek yeri belirleyin
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Excel Dosyası Kaydet",
                    FileName = "CariHesapEkstre.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // GridControl'ü Excel formatında dışa aktar
                    gC1.ExportToXlsx(saveFileDialog.FileName);
                    MessageBox.Show("Veriler başarıyla Excel'e aktarıldı.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel'e aktarma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void gridView1_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
        {
            if (e.MenuType == DevExpress.XtraGrid.Views.Grid.GridMenuType.Row)
            {
                var rowH = gridView1.FocusedRowHandle;
                var focusRowView = (DataRowView)gridView1.GetFocusedRow();
                if (focusRowView == null || focusRowView.IsNew) return;

                if (rowH >= 0)
                {
                    popupMenu1.ShowPopup(new Point(MousePosition.X, MousePosition.Y));
                }
                else
                {
                    popupMenu1.HidePopup();
                }
            }
        }
    }
}