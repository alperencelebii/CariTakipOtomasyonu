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
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\Data\db.mdb");
        public XtraAnaSayfa af;
        public int id;
        bool kaydet = true;
        PrintDocument printDocument = new PrintDocument();

        public XtraCariHareket()
        {
            InitializeComponent();
            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
            this.BtnYazdir.Click += new System.EventHandler(this.BtnYazdir_Click); // Butonun Click olayını bağla
        }

        private void BtnKapat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void bakiye_hesapla()
        {
            OleDbDataAdapter adp = new OleDbDataAdapter("Select * From TBLCARIHAREKET where CARI_ID=" + id, baglanti);
            DataTable da = new DataTable();
            da.Clear();
            adp.Fill(da);
            int son = da.Rows.Count;
            double borc = 0;
            double alacak = 0;
            double bakiye = 0;

            for (int i = 0; i < son; i++)
            {
                int kid = Convert.ToInt32(da.Rows[i]["ID"].ToString());
                borc = Convert.ToDouble(da.Rows[i]["BORC"].ToString());
                alacak = Convert.ToDouble(da.Rows[i]["ALACAK"].ToString());
                bakiye = bakiye + (borc - alacak);
                OleDbCommand kmt = new OleDbCommand("update TBLCARIHAREKET set BAKIYE = '" + bakiye + "' where ID=" + kid, baglanti);
                baglanti.Open();
                kmt.ExecuteNonQuery();
                baglanti.Close();
            }
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
                kmt.Parameters.AddWithValue("TARIH", DateTime.Now.ToShortDateString());
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

        void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                Font font = new Font("Arial", 8); // Smaller font
                Font titleFont = new Font("Arial", 12, FontStyle.Bold); // Slightly larger font for the title
                Font headerFont = new Font("Arial", 16, FontStyle.Bold); // Larger font for "Gökçe Tarım"
                float lineHeight = font.GetHeight() + 2; // Smaller line spacing
                float columnWidth = 80; // Narrower column width

                // Page width
                float pageWidth = e.PageBounds.Width;
                float y = 30; // Starting position from the top edge of the page

                // Page header
                string header = "Gökçe Tarım";
                float headerWidth = e.Graphics.MeasureString(header, headerFont).Width;
                e.Graphics.DrawString(header, headerFont, Brushes.Black, (pageWidth - headerWidth) / 2, y);
                y += headerFont.GetHeight() + 5;

                // Draw line below header
                e.Graphics.DrawLine(Pens.Black, 30, y, pageWidth - 30, y);
                y += 10;

                // Customer name and balance
                string title = LBCari.Text;
                string bakiye = LbBakiye.Text.Split(':')[1].Trim();
                string bakiyeText = "Bakiye: " + bakiye;
                float titleWidth = e.Graphics.MeasureString(title, titleFont).Width;
                float bakiyeWidth = e.Graphics.MeasureString(bakiyeText, font).Width;
                e.Graphics.DrawString(title, titleFont, Brushes.Black, (pageWidth - titleWidth) / 2, y);
                y += titleFont.GetHeight() + 5;
                e.Graphics.DrawString(bakiyeText, font, Brushes.Black, (pageWidth - bakiyeWidth) / 2, y);
                y += font.GetHeight() + 15;

                // Center the table
                float tableWidth = 7 * columnWidth; // Width of 7 columns
                float x = (pageWidth - tableWidth) / 2;

                // Header row
                e.Graphics.DrawString("Tarih", font, Brushes.Black, x, y);
                e.Graphics.DrawString("Stok Adı", font, Brushes.Black, x + columnWidth, y);
                e.Graphics.DrawString("Adet", font, Brushes.Black, x + 2 * columnWidth, y);
                e.Graphics.DrawString("Birim Fiyat", font, Brushes.Black, x + 3 * columnWidth, y);
                e.Graphics.DrawString("Toplam", font, Brushes.Black, x + 4 * columnWidth, y);
                e.Graphics.DrawString("Verilen", font, Brushes.Black, x + 5 * columnWidth, y);
                e.Graphics.DrawString("Bakiye", font, Brushes.Black, x + 6 * columnWidth, y);

                y += lineHeight;

                DataTable dataTable = (DataTable)gridControl1.DataSource;

                // Sort the DataTable by the "TARIH" column in ascending order
                DataView dataView = dataTable.DefaultView;
                dataView.Sort = "TARIH ASC";
                DataTable sortedDataTable = dataView.ToTable();

                foreach (DataRow row in sortedDataTable.Rows)
                {
                    if (y + lineHeight > e.MarginBounds.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    string tarih = Convert.ToDateTime(row["TARIH"]).ToShortDateString();
                    string stokAdi = row["STOK"].ToString();
                    string adet = row["ADET"].ToString();
                    string birimFiyat = row["BFIYAT"].ToString();
                    string borc = row["BORC"].ToString();
                    string alacak = row["ALACAK"].ToString();
                    string bakiyeRow = row["BAKIYE"].ToString();

                    e.Graphics.DrawString(tarih, font, Brushes.Black, x, y);
                    e.Graphics.DrawString(stokAdi, font, Brushes.Black, x + columnWidth, y);
                    e.Graphics.DrawString(adet, font, Brushes.Black, x + 2 * columnWidth, y);
                    e.Graphics.DrawString(birimFiyat, font, Brushes.Black, x + 3 * columnWidth, y);
                    e.Graphics.DrawString(borc, font, Brushes.Black, x + 4 * columnWidth, y);
                    e.Graphics.DrawString(alacak, font, Brushes.Black, x + 5 * columnWidth, y);
                    e.Graphics.DrawString(bakiyeRow, font, Brushes.Black, x + 6 * columnWidth, y);

                    y += lineHeight;
                }

                // Total balance row
                y += lineHeight;
                e.Graphics.DrawString("Toplam Bakiye:", font, Brushes.Black, x, y);
                e.Graphics.DrawString(bakiye, font, Brushes.Black, x + columnWidth, y);

                e.HasMorePages = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yazdırma işlemi sırasında bir hata oluştu: " + ex.Message);
            }
        }
    }
}