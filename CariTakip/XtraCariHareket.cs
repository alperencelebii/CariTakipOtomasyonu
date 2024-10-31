using DevExpress.ClipboardSource.SpreadsheetML;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CariTakip
{
    public partial class XtraCariHareket : Form
    {
        public OleDbConnection baglanti { get; set; }
        public XtraAnaSayfa af;
        public int id;
        bool kaydet = true;
        PrintDocument printDocument = new PrintDocument();

        private bool isDragging = false;
        private Point startPoint = new Point(0, 0);

        public XtraCariHareket()
        {
            InitializeComponent();
            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
            this.BtnYazdir.Click += new System.EventHandler(this.BtnYazdir_Click);

            datase.MouseDown += new MouseEventHandler(DateTimePicker_MouseDown);
            datase.MouseMove += new MouseEventHandler(DateTimePicker_MouseMove);
            datase.MouseUp += new MouseEventHandler(DateTimePicker_MouseUp);
        }

        private void DateTimePicker_MouseDown(object sender, MouseEventArgs e)
        {
            isDragging = true;
            startPoint = new Point(e.X, e.Y);
        }

        private void DateTimePicker_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point p = PointToClient(MousePosition);
                datase.Location = new Point(p.X - startPoint.X, p.Y - startPoint.Y);
            }
        }

        private void DateTimePicker_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }

        private void BtnKapat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void bakiye_hesapla()
        {
            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARIHAREKET where CARI_ID=" + id + " ORDER BY TARIH ASC", baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);

            double borc = 0;
            double alacak = 0;
            double bakiye = 0;

            for (int i = 0; i < da.Rows.Count; i++)
            {
                int kid = Convert.ToInt32(da.Rows[i]["ID"]);
                borc = Convert.ToDouble(da.Rows[i]["BORC"]);
                alacak = Convert.ToDouble(da.Rows[i]["ALACAK"]);
                bakiye += (borc - alacak);

                OleDbCommand kmt = new OleDbCommand("update TBLCARIHAREKET set BAKIYE = @BAKIYE where ID = @ID", baglanti);
                kmt.Parameters.AddWithValue("@BAKIYE", bakiye);
                kmt.Parameters.AddWithValue("@ID", kid);
                baglanti.Open();
                kmt.ExecuteNonQuery();
                baglanti.Close();
            }

            // Cari toplam bakiyeyi de güncelle
            UpdateCariBakiye(bakiye);
        }

        void UpdateCariBakiye(double bakiye)
        {
            OleDbCommand kmt = new OleDbCommand("update TBLCARI set BAKIYE = @BAKIYE where ID = @ID", baglanti);
            kmt.Parameters.AddWithValue("@BAKIYE", bakiye);
            kmt.Parameters.AddWithValue("@ID", id);
            baglanti.Open();
            kmt.ExecuteNonQuery();
            baglanti.Close();
        }

        void göster()
        {
            bakiye_hesapla();
            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARIHAREKET where CARI_ID=" + id + " order by TARIH desc ", baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);
            gridControl1.DataSource = da;

            adp = new OleDbDataAdapter("Select * From TBLCARI where ID=" + id, baglanti);
            da = new DataTable();
            da.Clear();
            adp.Fill(da);

            TxtStok.Text = "";
            TxtAdet.Text = "1";
            TxtBF.Text = "0";
            TxtTutar.Text = "0";

            LBCari.Text = da.Rows[0]["ADISOYADI"].ToString();
            LbBakiye.Text = "Bakiye: " + da.Rows[0]["BAKIYE"].ToString();
            kaydet = true;
            BtnTemizle.Enabled = false;
            af.göster();
        }

        void stok_getir()
        {
            TxtStok.Properties.Items.Clear();

            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLSTOK order by STOK", baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);
            int son = da.Rows.Count;
            for (int i = 0; i < son; i++)
            {
                TxtStok.Properties.Items.Add(da.Rows[i]["STOK"].ToString());
            }
        }

        private void XtraCariHareket_Load(object sender, EventArgs e)
        {
            stok_getir();
            göster();
        }

        private void TxtStok_SelectedIndexChanged(object sender, EventArgs e)
        {
            stok_seçimi();
        }

        void stok_seçimi()
        {
            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLSTOK WHERE STOK='" + TxtStok.Text + "'", baglanti);
                DataTable da = new DataTable();
                da.Clear();
                adp.Fill(da);

                string gg = da.Rows[0]["GELIRGIDER"].ToString();
                TxtBF.Text = da.Rows[0]["TUTAR"].ToString();
                if (gg == "GELIR")
                {
                    rbGelir.Checked = true;
                }
                else
                {
                    rbGider.Checked = true;
                }
            }
            catch
            {

            }

        }

        private void TxtAdet_TextChanged(object sender, EventArgs e)
        {
            hesapla();
        }

        private void TxtBF_EditValueChanged(object sender, EventArgs e)
        {
            hesapla();
        }

        void hesapla()
        {
            try
            {
                if (TxtAdet.Text == "")
                {
                    TxtAdet.Text = "1";
                }

                if (TxtBF.Text == "")
                {
                    TxtBF.Text = "0";
                }

                double adet = Convert.ToDouble(TxtAdet.Text);
                double bf = Convert.ToDouble(TxtBF.Text);
                double tutar = bf * adet;
                TxtTutar.Text = tutar.ToString();
            }
            catch
            {
                TxtTutar.Text = "0";
            }

        }

        private void TxtStok_EditValueChanged(object sender, EventArgs e)
        {
            stok_seçimi();
        }

        private void BtnKaydet_Click(object sender, EventArgs e)
        {
            double borc, alacak;
            borc = 0;
            alacak = 0;

            if (MessageBox.Show("Kaydetmek istiyor musunuz?", "Kaydet", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                OleDbCommand kmt;
                kmt = new OleDbCommand("insert into TBLCARIHAREKET(CARI_ID,TARIH,STOK,ADET,BFIYAT,BORC,ALACAK) VALUES(@CARI_ID,@TARIH,@STOK,@ADET,@BFIYAT,@BORC,@ALACAK)", baglanti);
                kmt.Parameters.AddWithValue("CARI_ID", id);

                // Seçilen tarihi kullanıyoruz
                kmt.Parameters.AddWithValue("TARIH", datase.Value.ToShortDateString());

                kmt.Parameters.AddWithValue("STOK", TxtStok.Text);
                kmt.Parameters.AddWithValue("ADET", TxtAdet.Text);
                kmt.Parameters.AddWithValue("BFIYAT", Convert.ToDouble(TxtBF.Text));
                if (rbGelir.Checked == true)
                {
                    alacak = Convert.ToDouble(TxtTutar.Text);
                    kmt.Parameters.AddWithValue("BORC", 0);
                    kmt.Parameters.AddWithValue("ALACAK", Convert.ToDouble(TxtTutar.Text));
                }
                else
                {
                    borc = Convert.ToDouble(TxtTutar.Text);
                    kmt.Parameters.AddWithValue("BORC", Convert.ToDouble(TxtTutar.Text));
                    kmt.Parameters.AddWithValue("ALACAK", 0);
                }
                baglanti.Open();
                kmt.ExecuteNonQuery();
                baglanti.Close();

                OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARI where ID=" + id, baglanti);
                DataTable da = new DataTable();
                da.Clear();
                adp.Fill(da);

                double bakiye = Convert.ToDouble(da.Rows[0]["BAKIYE"].ToString());
                bakiye = bakiye + (borc - alacak);
                kmt = new OleDbCommand("update TBLCARI set BAKIYE=@BAKIYE where ID=" + id, baglanti);
                kmt.Parameters.AddWithValue("BAKIYE", bakiye);
                baglanti.Open();
                kmt.ExecuteNonQuery();
                baglanti.Close();

                göster();
            }
        }

        private void BtnSil_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (MessageBox.Show("Kaydı silmek istiyor musunuz?", "Sil?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                sil();
            }
        }

        void sil()
        {
            double borc, alacak;
            borc = 0;
            alacak = 0;
            int ch_id = Convert.ToInt32(gridView1.GetFocusedRowCellValue("ID").ToString());
            borc = Convert.ToDouble(gridView1.GetFocusedRowCellValue("BORC").ToString());
            alacak = Convert.ToDouble(gridView1.GetFocusedRowCellValue("ALACAK").ToString());
            OleDbCommand kmt;
            kmt = new OleDbCommand("delete from TBLCARIHAREKET where ID=" + ch_id, baglanti);
            baglanti.Open();
            kmt.ExecuteNonQuery();
            baglanti.Close();

            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARI where ID=" + id, baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);
            double bakiye = Convert.ToDouble(da.Rows[0]["BAKIYE"].ToString());
            bakiye = bakiye - (borc - alacak);
            kmt = new OleDbCommand("update TBLCARI set BAKIYE=@BAKIYE where ID=" + id, baglanti);
            kmt.Parameters.AddWithValue("BAKIYE", bakiye);
            baglanti.Open();
            kmt.ExecuteNonQuery();
            baglanti.Close();

            göster();
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
                    popupMenu1.ShowPopup(barManager1, new Point(MousePosition.X, MousePosition.Y));
                }
                else
                {
                    popupMenu1.HidePopup();
                }
            }
        }

        private void BtnYazdir_Click(object sender, EventArgs e)
        {
            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = printDocument;
            printPreviewDialog.ShowDialog();
        }

        private int currentRow = 0; // Satır indeksi, sayfa boyunca ilerlemek için

        // Sınıf seviyesinde toplam değişkenleri tanımla
        decimal toplamSum = 0;
        decimal verilenSum = 0;

        void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                Font font = new Font("Arial", 8);
                Font titleFont = new Font("Arial", 12, FontStyle.Bold);
                Font headerFont = new Font("Arial", 16, FontStyle.Bold);
                Font groupFont = new Font("Arial", 10, FontStyle.Bold);

                float lineHeight = font.GetHeight() + 2;
                float columnWidth = 80;

                float pageWidth = e.PageBounds.Width;
                float y = 30;

                // Sayfa başlığı
                string header = "Cari Hesap Ekstre Raporu";
                float headerWidth = e.Graphics.MeasureString(header, headerFont).Width;
                e.Graphics.DrawString(header, headerFont, Brushes.Black, (pageWidth - headerWidth) / 2, y);
                y += headerFont.GetHeight() + 5;

                // Alt çizgi (Metin uzunluğunda ve daha kısa)
                float underlinePadding = 10;
                float underlineStartX = (pageWidth - headerWidth) / 2 + underlinePadding;
                float underlineEndX = (pageWidth + headerWidth) / 2 - underlinePadding;

                e.Graphics.DrawLine(Pens.Black, underlineStartX, y, underlineEndX, y);
                y += 10;

                // "Gökçe Tarım" başlığı için daha küçük bir font
                Font smallerFont = new Font(headerFont.FontFamily, headerFont.Size - 2, headerFont.Style);
                string header2 = "Gökçe Tarım";
                float header2Width = e.Graphics.MeasureString(header2, smallerFont).Width;
                e.Graphics.DrawString(header2, smallerFont, Brushes.Black, (pageWidth - header2Width) / 2, y);
                y += smallerFont.GetHeight() + 5;

                // Alt çizgi (Boydan boya)
                e.Graphics.DrawLine(Pens.Black, 30, y, pageWidth - 30, y);
                y += 10;

                // Mavi arka plan için genişletilmiş dikdörtgen
                RectangleF headerRect = new RectangleF(30, y, pageWidth - 60, 40);
                e.Graphics.FillRectangle(Brushes.SteelBlue, headerRect);

                // Cari ismini ortala
                string cariIsmi = LBCari.Text;
                float cariIsmiWidth = e.Graphics.MeasureString(cariIsmi, groupFont).Width;
                e.Graphics.DrawString(cariIsmi, groupFont, Brushes.White, (pageWidth - cariIsmiWidth) / 2, y + 5);

                y += 25; // Cari ismi ve tablo başlıkları arasına biraz boşluk bırakıyoruz

                // Tablo başlıkları (mavi dikdörtgenin içinde)
                e.Graphics.DrawString("Tarih", font, Brushes.White, 35, y);
                e.Graphics.DrawString("Stok Adı", font, Brushes.White, 35 + columnWidth, y);
                e.Graphics.DrawString("Adet", font, Brushes.White, 35 + 2 * columnWidth + 70, y);
                e.Graphics.DrawString("Birim Fiyat", font, Brushes.White, 35 + 3 * columnWidth + 60, y);
                e.Graphics.DrawString("Toplam", font, Brushes.White, 35 + 4 * columnWidth + 100, y);
                e.Graphics.DrawString("Verilen", font, Brushes.White, 35 + 5 * columnWidth + 120, y);
                e.Graphics.DrawString("Bakiye", font, Brushes.White, 35 + 6 * columnWidth + 130, y);

                y += headerRect.Height - 20; // Mavi dikdörtgenin altına geçiş yapıyoruz

                // Veriler burada hizalanacak
                DataTable dataTable = (DataTable)gridControl1.DataSource;

                DataView dataView = dataTable.DefaultView;
                dataView.Sort = "TARIH ASC, BAKIYE ASC";
                DataTable sortedDataTable = dataView.ToTable();

                StringFormat stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Far;

                // İlk sayfa mı kontrol edelim
                if (currentRow == 0)
                {
                    toplamSum = 0;
                    verilenSum = 0;
                }

                while (currentRow < sortedDataTable.Rows.Count)
                {
                    DataRow row = sortedDataTable.Rows[currentRow];

                    if (y + lineHeight > e.MarginBounds.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    string tarih = Convert.ToDateTime(row["TARIH"]).ToString("dd.MM.yyyy");
                    string stokAdi = row["STOK"].ToString();
                    string adet = Convert.ToInt32(row["ADET"]).ToString("N0");
                    string birimFiyat = Convert.ToDecimal(row["BFIYAT"]).ToString("N2");
                    decimal borc = Convert.ToDecimal(row["BORC"]);
                    decimal alacak = Convert.ToDecimal(row["ALACAK"]);
                    string bakiyeRow = Convert.ToDecimal(row["BAKIYE"]).ToString("N2");

                    toplamSum += borc;
                    verilenSum += alacak;

                    e.Graphics.DrawString(tarih, font, Brushes.Black, 35, y);
                    e.Graphics.DrawString(stokAdi, font, Brushes.Black, 35 + columnWidth, y);
                    e.Graphics.DrawString(adet, font, Brushes.Black, new RectangleF(35 + 2 * columnWidth + 20, y, columnWidth, lineHeight), stringFormat);
                    e.Graphics.DrawString(birimFiyat, font, Brushes.Black, new RectangleF(35 + 3 * columnWidth + 40, y, columnWidth, lineHeight), stringFormat);
                    e.Graphics.DrawString(borc.ToString("N2"), font, Brushes.Black, new RectangleF(35 + 4 * columnWidth + 60, y, columnWidth, lineHeight), stringFormat);
                    e.Graphics.DrawString(alacak.ToString("N2"), font, Brushes.Black, new RectangleF(35 + 5 * columnWidth + 80, y, columnWidth, lineHeight), stringFormat);
                    e.Graphics.DrawString(bakiyeRow, font, Brushes.Black, new RectangleF(35 + 6 * columnWidth + 100, y, columnWidth, lineHeight), stringFormat);

                    y += lineHeight;
                    currentRow++;
                }

                // Eğer son sayfadaysak, toplamları çiz
                if (!e.HasMorePages)
                {
                    y += lineHeight;
                    e.Graphics.DrawLine(Pens.Black, 30, y, pageWidth - 30, y); // Çizgi ekledik
                    y += 5;

                    float boxHeight = 30;
                    float boxWidth = pageWidth / 3 - 40;
                    float boxY = y + 20;  // Dikdörtgenler için alt konum

                    // Sol dikdörtgen (Borç)
                    e.Graphics.DrawString("BORÇ", font, Brushes.Black, 30, y);  // Başlığı daha küçük fontla yazıyoruz
                    RectangleF borcRect = new RectangleF(30, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, borcRect);
                    e.Graphics.DrawString(toplamSum.ToString("N2"), titleFont, Brushes.White, borcRect.X + 10, borcRect.Y + 5);

                    // Orta dikdörtgen (Verilen)
                    e.Graphics.DrawString("Verilen", font, Brushes.Black, borcRect.Right + 20, y);  // Başlığı daha küçük fontla yazıyoruz

                    RectangleF alacakRect = new RectangleF(borcRect.Right + 20, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, alacakRect);
                    e.Graphics.DrawString(verilenSum.ToString("N2"), titleFont, Brushes.White, alacakRect.X + 10, alacakRect.Y + 5);

                    // Sağ dikdörtgen (Toplam Bakiye)
                    string bakiyeText = LbBakiye.Text.Split(':')[1].Trim();
                    decimal bakiye = Convert.ToDecimal(bakiyeText);

                    e.Graphics.DrawString("TOPLAM", font, Brushes.Black, alacakRect.Right + 20, y);  // Başlığı daha küçük fontla yazıyoruz
                    RectangleF toplamRect = new RectangleF(alacakRect.Right + 20, boxY, boxWidth, boxHeight);
                    e.Graphics.FillRectangle(Brushes.SteelBlue, toplamRect);
                    e.Graphics.DrawString(bakiye.ToString("N2"), titleFont, Brushes.White, toplamRect.X + 10, toplamRect.Y + 5);

                    e.HasMorePages = false;
                    currentRow = 0; // Tüm sayfalar yazdırıldıktan sonra sıfırla
                }
            }
            catch (Exception ex)

            {
                MessageBox.Show("Yazdırma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }

    }
}