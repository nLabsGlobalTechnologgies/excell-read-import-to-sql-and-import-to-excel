using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcellProject
{
    public partial class Form1 : Form
    {
        public string dosyaYolu { get; set; }
        public string dosyaAdi { get; set; }
        public string dosyaUzantisi { get; set; }

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Sadece Excel dosyalarını göster
            openFileDialog.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dosyaYolu = openFileDialog.FileName; // Dosya yolunu alır
                dosyaAdi = Path.GetFileName(dosyaYolu); // Dosya adını alır
                dosyaUzantisi = Path.GetExtension(dosyaYolu); // Dosya uzantısını alır

                // Elde ettiğiniz dosya bilgilerini kullanarak bağlantı oluşturabilirsiniz
                string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'";

                OleDbConnection baglanti = new OleDbConnection(connString);
                baglanti.Open();

                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Products$]", baglanti);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt.DefaultView;

                baglanti.Close();

                // Dosya yolu, adı ve uzantısı burada kullanılabilir
                //MessageBox.Show("Seçilen Dosya Yolu: " + dosyaYolu + "\nDosya Adı: " + dosyaAdi + "\nDosya Uzantısı: " + dosyaUzantisi);
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            // DataGridView'daki verileri al
            DataView dataView = (DataView)dataGridView1.DataSource;
            DataTable dt = dataView.Table;

            // SQL Server veritabanına bağlantı oluştur
            string connectionString = "Data Source=.;Initial Catalog=TestDb;Integrated Security=True;";
            SqlConnection connection = new SqlConnection(connectionString);

            try
            {
                connection.Open();

                // Products tablosuna verileri eklemek için bir döngü oluştur
                foreach (DataRow row in dt.Rows)
                {
                    string insertQuery = "INSERT INTO Products (Name, Price) VALUES (@Name, @Price)";
                    SqlCommand cmd = new SqlCommand(insertQuery, connection);

                    // Parametreleri sorgu ile ilişkilendir
                    cmd.Parameters.AddWithValue("@Name", row[1]); // Name (1. sütun)
                    cmd.Parameters.AddWithValue("@Price", row[2]); // Price (2. sütun)

                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Veriler başarıyla SQL Server veritabanına ekledi.");
                SqlDataAdapter adapter = new SqlDataAdapter("select * from Products", connection);
                adapter.Fill(dt);
                dataGridView2.DataSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // DataGridView'daki verileri al
            DataView dataView = (DataView)dataGridView2.DataSource;
            DataTable dt = dataView.Table;

            // Bir SaveFileDialog oluştur
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Dosyası|*.xlsx"; // Sadece Excel dosyalarını göster

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string dosyaYolu = saveFileDialog.FileName;

                // Verileri Excel dosyasına kaydet
                using (var package = new ExcelPackage())
                {
                    var workSheet = package.Workbook.Worksheets.Add("Sayfa1"); // Sayfa adını istediğiniz gibi değiştirebilirsiniz

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            workSheet.Cells[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                        }
                    }

                    // Excel dosyasını kaydet
                    package.SaveAs(new FileInfo(dosyaYolu));
                    MessageBox.Show("Veriler başarıyla Excel dosyasına kaydedildi.");
                }
            }
        }

    }
}
