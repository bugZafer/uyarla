using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace uyarla
{
    public partial class fompk : Form
    {

        public fompk()
        {
            InitializeComponent();
        }

        private void verigetir()
        {
            varmi();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    if (xlRange != null)
                    {
                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;

                        int mlzKodIndex = -1, istenenTarihIndex = -1, miktarIndex = -1, kesinIndex = -1, spNo = -1;

                        for (int j = 1; j <= colCount; j++)
                        {
                            if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null)
                            {
                                if (xlRange.Cells[1, j].Value2.ToString() == "Malzeme No")
                                    mlzKodIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Sipariş Tarihi")
                                    istenenTarihIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Bakiye")
                                    miktarIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Sipariş Durumu")
                                    kesinIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Sipariş No")
                                    spNo = j;
                            }
                        }

                        if (mlzKodIndex == -1 || istenenTarihIndex == -1 || miktarIndex == -1 || kesinIndex == -1 || spNo == -1)
                        {
                            MessageBox.Show("Gerekli başlıklar bulunamadı.");
                            xlWorkbook.Close();
                            xlApp.Quit();
                            return;
                        }
                        for (int i = 2; i <= rowCount; i++)
                        {
                            if (xlRange.Cells[i, mlzKodIndex] != null && xlRange.Cells[i, mlzKodIndex].Value2 != null &&
                                xlRange.Cells[i, miktarIndex] != null && xlRange.Cells[i, miktarIndex].Value2 != null &&
                                xlRange.Cells[i, istenenTarihIndex] != null && xlRange.Cells[i, istenenTarihIndex].Value2 != null &&
                                xlRange.Cells[i, kesinIndex] != null && xlRange.Cells[i, kesinIndex].Value2 != null &&
                                xlRange.Cells[i, spNo] != null && xlRange.Cells[i, spNo].Value2 != null)
                            {
                                string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                                string ksnInd = xlRange.Cells[i, kesinIndex].Value2.ToString();
                                string sipNo = xlRange.Cells[i, spNo].Value2.ToString();
                                double serialDate = Convert.ToDouble(xlRange.Cells[i, istenenTarihIndex].Value2);
                                DateTime istenenTarih = ConvertFromSerialDate(serialDate);
                                string formattedDate = istenenTarih.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

                                // Excel hücresinin değerini doğrudan okuyarak kontrol et
                                object miktarValue = xlRange.Cells[i, miktarIndex].Value;
                                if (miktarValue != null)
                                {
                                    if (ksnInd == "Kesin")
                                    {
                                        dataGridView1.Rows.Add(mlzKod, miktarValue, formattedDate, sipNo);

                                    }
                                    else if (ksnInd == "Planli") //ÖNGÖRÜ OLAYI
                                    {
                                        dataGridView2.Rows.Add(mlzKod, miktarValue, formattedDate, sipNo);

                                    }
                                }
                            }
                        }

                        MessageBox.Show("Aktarım tamamlandı.");
                    }
                    if (xlRange != null)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);
                        xlWorkbook.Close(false, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(xlWorkbook);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }


        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F8)
            {
                try
                {

                    Excel.Application xlApp = new Excel.Application();
                    if (xlApp == null)
                    {
                        // Excel yüklü değilse hata işlemleri
                        return;
                    }

                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    // Hücre başlıklarını Excel'e ekle
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        xlWorkSheet.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText;
                    }

                    // DataGridView'den verileri al ve Excel'e aktar
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++) // Subtract 1 to skip the last row
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            xlWorkSheet.Cells[i + 2, j + 1] = "'" + (dataGridView1.Rows[i].Cells[j].Value != null ? dataGridView1.Rows[i].Cells[j].Value.ToString() : "");
                        }
                    }

                    MessageBox.Show("Kesin Sipariş Aktarımı Bitti.");
                    xlApp.Visible = true;
                    // Excel nesnelerini serbest bırak
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu: " + ex.Message);
                }



            }
            else if (e.KeyCode == Keys.F5)
            {
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
            else if ((e.KeyCode == Keys.F10))
            {
                try
                {

                    Excel.Application xlApp = new Excel.Application();
                    if (xlApp == null)
                    {
                        // Excel yüklü değilse hata işlemleri
                        return;
                    }

                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    // Hücre başlıklarını Excel'e ekle
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        xlWorkSheet.Cells[1, j + 1] = dataGridView2.Columns[j].HeaderText;
                    }

                    // DataGridView'den verileri al ve Excel'e aktar
                    for (int i = 0; i < dataGridView2.Rows.Count - 1; i++) // Subtract 1 to skip the last row
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            xlWorkSheet.Cells[i + 2, j + 1] = "'" + (dataGridView2.Rows[i].Cells[j].Value != null ? dataGridView2.Rows[i].Cells[j].Value.ToString() : "");
                        }
                    }

                    MessageBox.Show("Öngörü Sipariş Aktarımı Bitti.");
                    xlApp.Visible = true;
                    // Excel nesnelerini serbest bırak
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu: " + ex.Message);
                }


            }

        }
       





















        private void ficosa()
        {
            varsasil();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int firma = 0;
                int bakiye = 0;
                int sText = 0;
                int kontrol = 0;
                string dateValue = "";
                string dateValue2 = "";
                string datevalueYeni = "";

                foreach (Excel.Range row in excelRange.Rows)
                {
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[2, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Referans")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }

                        kontrol++;
                    }
                    for (int jx = 1; jx <= excelRange.Columns.Count; jx++)
                    {
                        string cellValueg = excelRange.Cells[2, jx].Value?.ToString();
                        if (cellValueg != null && cellValueg == "Bakiye")
                        {
                            bakiye = jx;
                            kontrol = 1;
                            break;
                        }
                    }
                    for (int yl = 1; yl <= excelRange.Columns.Count; yl++)
                    {
                        string cellValueS = excelRange.Cells[2, yl].Value?.ToString();
                        if (cellValueS != null && cellValueS == "S")
                        {
                            kontrol = 1;

                            // S'nin bulunduğu sütunun altındaki satırları kontrol etmek için bir iç içe döngü
                            for (int i = 3; i <= excelRange.Rows.Count; i++)
                            {
                                string cellValueBelow = excelRange.Cells[i, yl].Value?.ToString();

                                // '*' karakterini ara
                                if (cellValueBelow != null && cellValueBelow.Contains("*"))
                                {
                                    sText = yl;
                                    kontrol = 1;
                                    // '*' bulunduğunda gerekli işlemleri yapabilirsiniz
                                    // İsterseniz döngüden çıkabilirsiniz
                                    break;
                                }
                            }

                            // Dış döngüden çıkabilirsiniz, eğer sadece ilk 'S' bulunan sütunu kontrol etmek istiyorsanız
                        }
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }


                    Dictionary<string, double> firmTotalMap = new Dictionary<string, double>();
                    Dictionary<string, double> firmTotalMap2 = new Dictionary<string, double>();
                    int startingColumn = Math.Min(bakiye, sText);
                    int endingColumn = Math.Max(bakiye, sText);
                    for (int j = startingColumn; j <= endingColumn; j++)
                    {
                        for (int i = 2; i <= excelRange.Rows.Count; i++)
                        {
                            object cellValue = excelRange.Cells[i, j].Value;
                            // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                            if (cellValue != null && double.TryParse(cellValue.ToString(), out double numericValue))
                            {
                                // Numeric değeri kullanabilirsiniz
                              
                                if (numericValue != 0)
                                {
                                        dateValue = excelRange.Cells[2, j].Value?.ToString();
                                   
                                        string firmaDatax = excelRange.Cells[i, firma].Value.ToString();

                                        if (firmaDatax == "3M51-R43404-A ")
                                        {
                                            firmaDatax = "3M51-R43404-A M6";
                                        }

                                        // Eğer firmaDatax bilgisi varsa, firmaya ait toplamı hesapla ve yaz
                                        if (!string.IsNullOrEmpty(firmaDatax))
                                        {
                                            if (!firmTotalMap.ContainsKey(firmaDatax))
                                            {
                                                firmTotalMap[firmaDatax] = 0; // Firma için toplamı sıfırla
                                            }

                                            firmTotalMap[firmaDatax] += numericValue;
                                        }
                                }
                            }
                        }
                    }

                    for (int j2 = endingColumn; j2 <= excelRange.Columns.Count; j2++)  // ÖNGÖRÜ OLAYI
                    {
                        for (int i2 = 2; i2 <= excelRange.Rows.Count; i2++)
                        {
                            object cellValue2 = excelRange.Cells[i2, j2].Value;
                            // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                            if (cellValue2 != null && double.TryParse(cellValue2.ToString(), out double numericValue2))
                            {
                                // Numeric değeri kullanabilirsiniz

                                if (numericValue2 != 0)
                                {
                                    dateValue2 = excelRange.Cells[2, j2].Value?.ToString();
                                    if (DateTime.TryParseExact(dateValue2, "d.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                    {
                                        datevalueYeni = parsedDate.ToString("dd.MM.yyyy");
                                    }
                                    string firmaDatax2 = excelRange.Cells[i2, firma].Value.ToString();

                                    if (firmaDatax2 == "3M51-R43404-A ")
                                    {
                                        firmaDatax2 = "3M51-R43404-A M6";
                                    }

                                    // Eğer firmaDatax bilgisi varsa, firmaya ait toplamı hesapla ve yaz
                                    if (!string.IsNullOrEmpty(firmaDatax2))
                                    {
                                        if (!firmTotalMap2.ContainsKey(firmaDatax2))
                                        {
                                            firmTotalMap2[firmaDatax2] = 0; // Firma için toplamı sıfırla
                                        }

                                        firmTotalMap2[firmaDatax2] += numericValue2;
                                    }
                                }
                            }
                        }
                    }
                    
                    // Hesaplanan toplamları DataGridView'e ekleyin
                    foreach (var kvp in firmTotalMap)
                    {
                        string firmaDatax = kvp.Key;
                        double totalValue = kvp.Value;
                        
                        if (DateTime.TryParseExact(dateValue, "d.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                        {
                            DateTime previousSaturday = parsedDate.Date.AddDays(-(int)parsedDate.DayOfWeek - 1);
                            dataGridView1.Rows.Add(firmaDatax, totalValue, previousSaturday.ToString("dd.MM.yyyy"));
                        }
                        else if (DateTime.TryParseExact(dateValue, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDateAlt))
                        {
                            DateTime previousSaturday = parsedDateAlt.Date.AddDays(-(int)parsedDateAlt.DayOfWeek - 1);
                            dataGridView1.Rows.Add(firmaDatax, totalValue, previousSaturday.ToString("dd.MM.yyyy"));
                        }
                        else
                        {
                            // dateValue, beklenen formatta bir tarih değilse, bir hata mesajı gösterilebilir.
                            MessageBox.Show($"Hata: Geçerli bir tarih bulunamadı. Orijinal Tarihi: {dateValue}");
                        }
                    }

                    foreach (var kvp2 in firmTotalMap2)
                    {
                        string firmaDatax2 = kvp2.Key;
                        double totalValue2 = kvp2.Value;
                     
                        // Dönüşüm başarılı oldu, datevalueYeni şu anda doğru bir tarih değerini içeriyor
                        dataGridView2.Rows.Add(firmaDatax2, totalValue2, datevalueYeni);

                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    break;

                }


                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            ficosa();
        }

        private DateTime ConvertFromSerialDate(double serialDate)
        {
            if (serialDate > 0)
            {
                DateTime dateTime = new DateTime(1899, 12, 30).AddDays(serialDate);
                return dateTime;
            }
            return DateTime.MinValue;
        }


        int firmaBilgi;
       


        private void FindColumnBackgroundColor()
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 6;
                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Ürün No") // Değiştirmeniz gerekebilir
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }

                        kontrol++;

                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;

                    }
                    Dictionary<string, double> yesilMiktarlar = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();
                    int ilkYesilSutunNumarasi = -1;

                    int startingColumn = bakiye; // Bakiye hücresinin sağında(+1) tarih var ordan itibaren al
                    // Yeşil renkli hücre işlemleri // ÖNGÖRÜ İŞLEMLERİ
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 147, 255, 149)))
                        {
                            ilkYesilSutunNumarasi = cell.Column;

                            for (int j = ilkYesilSutunNumarasi; j <= excelRange.Columns.Count; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;

                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[row.Row, firmaBilgi].Value.ToString();
                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        // Veri daha önce eklenmemişse ekle ve geçici listeyi güncelle
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                            if (firmaDatax == "3A00 217 008-01")
                                            {
                                                firmaDatax = "3A00 217 008";
                                            }

                                            // DataGridView'e ekle, bu sefer format dönüşümü yaparak ekleyin
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Diğer renkli hücre işlemleri
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 180, 214, 238)) ||
                            cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 174, 137)))
                        {

                            for (int j = startingColumn; j <= ilkYesilSutunNumarasi - 1; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;

                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();

                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");

                                        // Veri daha önce eklenmemişse ekle ve geçici listeyi güncelle
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                            if (firmaDatax == "3A00 217 008-01")
                                            {
                                                firmaDatax = "3A00 217 008";
                                            }

                                            // DataGridView'e ekle, bu sefer format dönüşümü yaparak ekleyin
                                            dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }



                }

                MessageBox.Show("Aktarım Tamamlandı.");
                excelWorkbook.Close();
                excelApp.Quit();

            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //arama yapmak için text giriyoruz
            string Aratxt = textBox1.Text.Trim().ToUpper();
            int j = -1;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                foreach (DataGridViewCell cell in dataGridView1.Rows[i].Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToUpper().Contains(Aratxt))
                    {
                        cell.Style.BackColor = Color.Yellow;
                        j = 0;
                    }
                }
            }

            if (j == -1)
            {
                MessageBox.Show("Kayıt bulunamadı!");
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
          
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            FindColumnBackgroundColor();
        }


        private void teklas()
        {
            varmi();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    if (xlRange != null)
                    {
                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;
                        string ongDate = "";

                        List<string> headers = new List<string>();
                        List<string> headerF = new List<string>();
                        for (int j = 1; j <= colCount; j++)
                        {
                            if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null)
                            {
                                headers.Add(xlRange.Cells[1, j].Value2.ToString());
                                headerF.Add(xlRange.Cells[1, j].Value2.ToString());
                            }
                        }
                        if (headers.Contains("customerItem") && headers.Contains("qtyOrdered") && headers.Contains("dueDate") ||
                            headerF.Contains("customerItem") && headerF.Contains("qtyOrdered") && headerF.Contains("dueDate") && headerF.Contains("stat"))

                        {
                            for (int i = 2; i <= rowCount; i++)
                            {
                                string STAT = xlRange.Cells[i, headers.IndexOf("stat") + 1].Value2.ToString();
                                if (STAT == "O".ToString())
                                {
                                    string customerItem = xlRange.Cells[i, headers.IndexOf("customerItem") + 1].Value2.ToString();
                                    string qtyOrdered = xlRange.Cells[i, headers.IndexOf("qtyOrdered") + 1].Value2.ToString();
                                    string OrderNo = xlRange.Cells[i, headers.IndexOf("purchaseOrderNumber") + 1].Value2.ToString();
                                    string houseId = xlRange.Cells[i, headers.IndexOf("deliveryWarehouse") + 1].Value2.ToString();
                                    string excelDateStr = xlRange.Cells[i, headers.IndexOf("dueDate") + 1].Value2.ToString();
                                    DateTime dueDate;
                                    if (DateTime.TryParseExact(excelDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dueDate))
                                    {
                                        ongDate = dueDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                                        // Diğer işlemler...
                                    }
                                    else
                                    {
                                        double excelDateNum = double.Parse(excelDateStr);
                                        dueDate = new DateTime(1900, 1, 1).AddDays(excelDateNum - 2); // -2 çünkü 1900 tarihinde 1 Ocak olarak başlıyor ve seri numarası 2 gün eksik
                                        ongDate = dueDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                                    }



                                    if (houseId == "KOMPONENT DEPO - Bulgaria KOMPONENT DEPO")
                                    {
                                        houseId = "Bulgaria";
                                    }
                                    else if (houseId== "KOMPONENT DEPO -  GOSB1")
                                    {
                                        houseId = "GOSB1";
                                    }
                                    else if (houseId == "BARTIN-2 -  BARTIN-2")
                                    {
                                        houseId = "BARTIN-2";
                                    }
                                    else if (houseId== "MUALLIMKOY -  MUALLIMKOY")
                                    {
                                        houseId= "MUALLIMKOY";
                                    }


                                    dataGridView1.Rows.Add(customerItem, qtyOrdered, ongDate, OrderNo + " - " + houseId);

                                }
                                else if (STAT == "F".ToString()) //ÖNGÖRÜ OLAYI
                                {
                                    string customerItem = xlRange.Cells[i, headerF.IndexOf("customerItem") + 1].Value2.ToString();
                                    string qtyOrdered = xlRange.Cells[i, headerF.IndexOf("qtyOrdered") + 1].Value2.ToString();
                                    string excelDateStr = xlRange.Cells[i, headers.IndexOf("dueDate") + 1].Value2.ToString();
                                    DateTime dueDate;
                                    if (DateTime.TryParseExact(excelDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dueDate))
                                    {
                                        ongDate = dueDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                                        // Diğer işlemler...
                                    }
                                    else
                                    {
                                        double excelDateNum = double.Parse(excelDateStr);
                                        dueDate = new DateTime(1900, 1, 1).AddDays(excelDateNum - 2); // -2 çünkü 1900 tarihinde 1 Ocak olarak başlıyor ve seri numarası 2 gün eksik
                                        ongDate = dueDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                                    }
                                  
                                    dataGridView2.Rows.Add(customerItem, qtyOrdered, ongDate);
                                }
                            }
                            MessageBox.Show("Aktarım Tamamlandı.");
                        }
                        else
                        {
                            MessageBox.Show("Gerekli başlıklar bulunamadı.", "Başlık Kontrolü");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Dosya boş.");
                    }

                    xlWorkbook.Close(false);
                    xlApp.Quit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }

        }



        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            teklas();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            firmaBilgi = 0;
            FindColumnBackgroundColor();

        }
        private void aka()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 0;
                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        if (excelRange.Cells[j, 1].Value2 != null && excelRange.Cells[j, 1].Value2.ToString() == "1")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }

                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string tarihValue = excelRange.Cells[6, j].Value2 != null ? excelRange.Cells[6, j].Value2.ToString() : "";

                        if (tarihValue != null && tarihValue.Length == 8 &&
                                    tarihValue[2] == '.' && tarihValue[5] == '.')
                        {
                            bakiye = j;
                            kontrol = 1;
                            break;

                        }
                        kontrol++;
                    }
                    int gercek = 0;
                    for (int jg = 1; jg <= excelRange.Columns.Count; jg++)
                    {
                        if (excelRange.Cells[6, jg].Value2 != null && excelRange.Cells[6, jg].Value2.ToString() == "Bakiye")
                        {
                            gercek = jg;
                            kontrol = 1;
                            break;
                        }
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }
                    Dictionary<string, double> firmTotalMapm = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();
                    HashSet<string> eklenenVeriler2 = new HashSet<string>();

                    DateTime today = DateTime.Now;
                    int daysUntilNextMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
                    // Eğer bugün Pazartesi ise 2 hafta sonrasını al, değilse bir sonraki Pazartesi'yi al
                    DateTime nextMonday = today.AddDays(daysUntilNextMonday == 0 ? 14 : daysUntilNextMonday + 7);

                    DateTime firstSaturday = today;

                    // Bugün Cumartesi değilse, bir sonraki Cumartesi'yi bul
                    if (today.DayOfWeek != DayOfWeek.Saturday)
                    {
                        int daysUntilNextSaturday = ((int)DayOfWeek.Saturday - (int)today.DayOfWeek + 7) % 7;
                        firstSaturday = today.AddDays(daysUntilNextSaturday);
                    }

                    // Diğer renkli hücre işlemleri

                    int startingColumn = bakiye;
                    int intValue=0;


                    for (int j = startingColumn; j <= excelWorksheet.UsedRange.Columns.Count; j++)
                    {
                        // En üst sütundaki tarih değerini al
                        string dateValue = excelRange.Cells[6, j].Value?.ToString();
                       
                        // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                        if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                        {
                            // İlgili sütunu kontrol et ve işlemleri yap
                            for (int i = firma; i <= excelWorksheet.UsedRange.Rows.Count; i++)
                            {
                                object cellValuex = excelRange.Cells[i, j].Value;

                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue))
                                {
                                    // Numeric değeri kullanabilirsiniz
                                    if (numericValue != 0)
                                    {
                                        object cellValue = excelRange.Cells[i, gercek].Value2;
                                        if (cellValue != null)
                                        {
                                            if (cellValue is double)
                                            {
                                                double doubleValue = (double)cellValue;
                                                intValue = (int)doubleValue;

                                                // intValue'i kullanabilirsiniz
                                            }
                                            else
                                            {
                                                // Hücrede sayısal bir değer yoksa veya double'a dönüştürülemezse bir işlem yapabilirsiniz.
                                            }
                                        }
                                         numericValue += intValue;
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[i, 2].Value.ToString();

                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        if (firmaDatax== "PZ31V04545APIA14")
                                        {
                                            firmaDatax = "PZ31-V04545-A-PIA-14-02";
                                        }
                                        else if (firmaDatax== "1290006160")
                                        {
                                            firmaDatax = "12900-06160";
                                        }
                                        else if (firmaDatax == "1290006200")
                                        {
                                            firmaDatax = "12900-06200";

                                        }
                                        else if (firmaDatax == "9652420008")
                                        {
                                            firmaDatax = "R965242008";
                                        }
                                        else if (firmaDatax == "9652420010")
                                        {
                                            firmaDatax = "R9652420010";
                                        }
                                        // Eğer tarih nextMonday'den küçükse, dataGridView1'e ekle, değilse dataGridView2'ye ekle
                                        if (firmTotalMapm.ContainsKey(firmaDatax))
                                        {
                                            firmTotalMapm[firmaDatax] += numericValue;
                                        }
                                        else
                                        {
                                            firmTotalMapm.Add(firmaDatax, numericValue);
                                        }
                                        if (DateTime.Compare(date, nextMonday) < 0)
                                        {
                                            dataGridView1.Rows.Clear(); // Önce dataGridView1'i temizle
                                            foreach (var kvp in firmTotalMapm)
                                            {
                                              
                                                dataGridView1.Rows.Add(kvp.Key, kvp.Value, firstSaturday.ToString("dd.MM.yyyy"));
                                            }
                                        }
                                        else
                                        {
                                            // ÖNGÖRÜ OLAYI
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, nextMonday.ToString("dd.MM.yyyy"));
                                        }

                                    }
                                }
                            }
                        }
                    }
                    break;



                }
                MessageBox.Show("Aktarım Tamamlandı.");

                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }


        private void button4_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            aka();
        }

        private void alkor()
        {

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int mlzKodIndex = -1;

                    // 2. satırdan "ÜRÜN KODU" başlığını bul
                    for (int col = 1; col <= xlRange.Columns.Count; col++)
                    {
                        if (xlRange.Cells[2, col].Value2 != null && xlRange.Cells[2, col].Value2.ToString() == "ÜRÜN KODU")
                        {
                            mlzKodIndex = col;
                            break;
                        }
                    }

                    if (mlzKodIndex == -1)
                    {
                        MessageBox.Show("Gerekli başlık bulunamadı.");
                        xlWorkbook.Close();
                        xlApp.Quit();
                        return;
                    }

                    // DataGridView'e sütun başlıklarını ekle

                    int rowCount = xlRange.Rows.Count;
                    string kodx = null;
                    DateTime ayBilgi = DateTime.MinValue;
                    string duzeltme = "Y06 0001335";
                    string duzeltme2 = "Y06 0001394";
                    // "ÜRÜN KODU" bulunduktan sonra altındaki verileri DataGridView'e ekleyin
                    for (int i = 3; i <= rowCount; i++) // 2. satırı atladık
                    {
                        if (xlRange.Cells[i, mlzKodIndex] != null && xlRange.Cells[i, mlzKodIndex].Value2 != null)
                        {
                            string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                            // "OCAK", "ŞUBAT", ..., "ARALIK" başlıklarını içeren ay kontrolü
                            bool ayVar = false;
                            for (int j = mlzKodIndex + 1; j <= xlRange.Columns.Count; j++)
                            {
                                if (xlRange.Cells[2, j] != null && xlRange.Cells[2, j].Value2 != null)
                                {
                                    string ay = xlRange.Cells[2, j].Value2.ToString();
                                    string veri = xlRange.Cells[i, j].Value2 != null ? xlRange.Cells[i, j].Value2.ToString() : "";
                                    if (mlzKod == duzeltme)
                                    {
                                        mlzKod = "Y060001335";
                                    }
                                    else if (mlzKod == duzeltme2)
                                    {
                                        mlzKod = "Y060001394";
                                    }
                                    else
                                    {


                                        // Eğer Excel sayfasındaki başlık içerisinde ay ifadesi varsa, DataGridView'e ekle
                                        if (ay.Contains("OCAK") || ay.Contains("ŞUBAT") || ay.Contains("MART") || ay.Contains("NİSAN") || ay.Contains("MAYIS") || ay.Contains("HAZİRAN") || ay.Contains("TEMMUZ") || ay.Contains("AĞUSTOS") || ay.Contains("EYLÜL") || ay.Contains("EKİM") || ay.Contains("KASIM") || ay.Contains("ARALIK"))
                                        {
                                            if (ay.Contains("OCAK"))
                                            {
                                                ayBilgi = new DateTime(2023, 01, 01);
                                            }
                                            else if (ay.Contains("ŞUBAT"))
                                            {
                                                ayBilgi = new DateTime(2023, 02, 01);

                                            }
                                            else if (ay.Contains("MART"))
                                            {
                                                ayBilgi = new DateTime(2023, 03, 01);

                                            }
                                            else if (ay.Contains("NİSAN"))
                                            {
                                                ayBilgi = new DateTime(2023, 04, 01);

                                            }
                                            else if (ay.Contains("MAYIS"))
                                            {
                                                ayBilgi = new DateTime(2023, 05, 01);

                                            }
                                            else if (ay.Contains("HAZİRAN"))
                                            {
                                                ayBilgi = new DateTime(2023, 06, 01);
                                            }
                                            else if (ay.Contains("TEMMUZ"))
                                            {
                                                ayBilgi = new DateTime(2023, 07, 01);

                                            }
                                            else if (ay.Contains("AĞUSTOS"))
                                            {
                                                ayBilgi = new DateTime(2023, 08, 01);

                                            }
                                            else if (ay.Contains("EYLÜL"))
                                            {
                                                ayBilgi = new DateTime(2023, 09, 01);

                                            }
                                            else if (ay.Contains("EKİM"))
                                            {
                                                ayBilgi = new DateTime(2023, 10, 01);

                                            }
                                            else if (ay.Contains("KASIM"))
                                            {
                                                ayBilgi = new DateTime(2023, 11, 01);

                                            }
                                            else if (ay.Contains("ARALIK"))
                                            {
                                                ayBilgi = new DateTime(2023, 12, 01);

                                            }
                                            kodx = veri;
                                            //    MessageBox.Show(kodx);

                                            if (int.TryParse(kodx, out int mlzKodInt) && mlzKodInt != 0)
                                            {

                                                ayVar = true;
                                                dataGridView1.Rows.Add(mlzKod, veri, ayBilgi.ToString("dd.MM.yyyy"));

                                            }
                                        }
                                    }
                                }
                            }

                            // Eğer hiç ay bulunmazsa, sadece "MlzKod" ekleyin
                            if (!ayVar)
                            {
                                dataGridView1.Rows.Add(mlzKod, "");
                            }
                        }
                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    if (xlRange != null)
                    {
                        Marshal.ReleaseComObject(xlRange);
                    }

                    if (xlWorksheet != null)
                    {
                        Marshal.ReleaseComObject(xlWorksheet);
                    }

                    xlWorkbook.Close(false, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(xlWorkbook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }


        private void alkorr_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            alkor();
        }
        private void kaplam()
        {
            varmi();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    bool müsteriNoFound = false;
                    // Excel sayfasındaki tüm hücreleri döngüye al
                    foreach (Excel.Range cell in xlRange)
                    {
                        if (cell.Value != null && cell.Value.ToString() == "Müşteri No")
                        {
                            müsteriNoFound = true;

                            int columnIndex = cell.Column;
                            string miktarValue = "";
                            string dortuncuSatirDegeri = "";
                            int rowIndex = cell.Row + 1;
                            while (xlWorksheet.Cells[rowIndex, columnIndex].Value != null)
                            {
                                object customerNoValue = xlRange.Range[cell.Address].Offset[rowIndex - cell.Row, 0].Value;
                                // Sadece gizli olmayan sütunlarda Üstteki Değer'i kontrol et
                                for (int targetColumnIndex = 1; targetColumnIndex <= xlWorksheet.UsedRange.Columns.Count; targetColumnIndex++)
                                {
                                    if (!xlWorksheet.UsedRange.Columns[targetColumnIndex].Hidden)
                                    {
                                        object usttekiDegerObj = xlWorksheet.Cells[1, targetColumnIndex].Value;
                                        string usttekiDeger = usttekiDegerObj?.ToString();
                                        miktarValue = xlWorksheet.Cells[rowIndex, targetColumnIndex].Value?.ToString();
                                        object dortuncuSatirDegeriObj = xlWorksheet.Cells[4, targetColumnIndex].Value;
                                        dortuncuSatirDegeri = null;
                                        if (miktarValue != "0")
                                        {
                                            if (dortuncuSatirDegeriObj != null && DateTime.TryParse(dortuncuSatirDegeriObj.ToString(), out DateTime dateValue))
                                            {
                                                dortuncuSatirDegeri = dateValue.ToString("dd.MM.yyyy");

                                                // Üstteki Değer dolu ve hedef sütununun altındaki Miktar değeri "0" değilse işlemleri gerçekleştir
                                                if (!string.IsNullOrEmpty(usttekiDeger))
                                                {
                                                    dataGridView1.Rows.Add(customerNoValue.ToString(), miktarValue, dortuncuSatirDegeri, usttekiDeger);
                                                }
                                                else
                                                {
                                                    dataGridView2.Rows.Add(customerNoValue.ToString(), miktarValue, dortuncuSatirDegeri, usttekiDeger);
                                                }
                                            }
                                        }
                                    }
                                }

                                rowIndex++;
                            }
                        }
                    }
                   

                    MessageBox.Show("Aktarım Tamamlandı.");
                    if (!müsteriNoFound)
                    {
                        MessageBox.Show("Müşteri No bulunamadı.");
                    }


                    // Excel uygulamasını kapatın
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // Kullanılan Excel nesnelerini serbest bırakın
                    ReleaseObject(xlRange);
                    ReleaseObject(xlWorksheet);
                    ReleaseObject(xlWorkbook);
                    ReleaseObject(xlApp);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        
        }
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Marshal.ReleaseComObject yöntemi çağrılamadı. Hata: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

       private void button1_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            kaplam();

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            varsasil();
            firmaBilgi = 3;
            FindColumnBackgroundColor();
            
        }

        private void opsn_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;

                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel.Range xlRange = null;

                try
                {
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(filePath);
                    xlWorksheet = xlWorkbook.Sheets[1];

                    // İlk dolu satırın indeksini bul
                    int firstDataRowIndex = FindFirstDataRowIndex(xlWorksheet);

                    if (firstDataRowIndex > 0)
                    {
                        // Var olan DataGridView kontrolüne ekle
                        for (int i = firstDataRowIndex; i <= xlWorksheet.UsedRange.Rows.Count; i++)
                        {
                            xlRange = xlWorksheet.Rows[i];

                            // İlk hücrenin değerini "MlzKod" sütununa ekle
                            string mlzKodValue = xlRange.Cells[1, 1].Text;

                            // Diğer hücreleri "Miktar" sütununa ekle
                            for (int j = 2; j <= xlWorksheet.UsedRange.Columns.Count; j++)
                            {
                                string miktarValue = xlRange.Cells[1, j].Text;
                                // İlgili tarih sütununu belirle
                                string tarihColumnName = xlWorksheet.Cells[4, j].Text;
                                string taih = xlWorksheet.Cells[4, j].Text;

                                // Noktaları temizle
                                miktarValue = miktarValue.Replace(".", "");

                                if (mlzKodValue == "113-01-70055724")
                                {
                                    mlzKodValue = "70055724";
                                }
                                else if (mlzKodValue == "113-01-PZ31-K624B93AP10A")
                                {
                                    mlzKodValue = "PZ31-K624B93-A-PIA-10";
                                }
                                else if (mlzKodValue == "113-06-W720476")
                                {
                                    mlzKodValue = "W720476";
                                }
                                else if (mlzKodValue == "113-01-H1BB-109A26APIA11")
                                {
                                    mlzKodValue = "H1BB-109A26-APIA-11";
                                }
                                if (miktarValue != "" && taih != "MOQ" && taih != null && taih != "")
                                {
                                    dataGridView1.Rows.Add(mlzKodValue, miktarValue, taih);
                                }
                            }
                        }

                        MessageBox.Show("Aktarım Tamamlandı.");
                    }
                    else
                    {
                        MessageBox.Show("Dolu satır bulunamadı.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Excel işlemleri sırasında bir hata oluştu: " + ex.Message);
                }
                if (xlApp != null)
                {
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // Kullanılan Excel nesnelerini serbest bırak
                    ReleaseObject(xlWorksheet);
                    ReleaseObject(xlWorkbook);
                    ReleaseObject(xlApp);

                    // Garbage Collector'ı çağır
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

        }
        private static int FindFirstDataRowIndex(Excel._Worksheet worksheet)
        {
            // İlk dolu satırın indeksini bul
            for (int i = 5; i <= worksheet.Rows.Count; i++)
            {
                Excel.Range row = worksheet.Rows[i];
                if (!string.IsNullOrWhiteSpace(row.Cells[1, 1].Text))
                {
                    return i;
                }
            }
            return -1; // Dolu satır bulunamazsa -1 döndür
        }

        private void fompak()
        {
            varmi();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string dosyaYolu = openFileDialog1.FileName;

                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel.Range xlRange = null;

                try
                {
                    // Excel uygulamasını oluştur
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(dosyaYolu);
                    // İlk sayfa ile çalıştığınızı varsayalım
                    xlWorksheet = xlWorkbook.Sheets[1];
                    xlRange = xlWorksheet.UsedRange;
                    // Başlangıç hücresini bul
                    int baslangicSatiri = -1;
                    int baslangicSutunu = -1;
                    int satirSayisi = xlRange.Rows.Count;
                    int sutunSayisi = xlRange.Columns.Count;
                    // "Referans", "İrsaliye Tarihi" ve "Bakiye" sütunlarını bul
                    int irsaliyeTarihiSutunu = -1;
                    int bakiyeSutunu = -1;
                    int sipNo = -1;
 
                    for (int j = 1; j <= sutunSayisi; j++)
                    {
                        if (xlRange.Cells[1, j].Value2 != null)
                        {
                            string baslik = xlRange.Cells[1, j].Value2.ToString();
                            if (baslik == "Referans")
                            {
                                baslangicSutunu = j;
                                baslangicSatiri = 2; // Başlangıç satırını bir sonraki satır olarak ayarla
                            }
                            else if (baslik == "İrsaliye Tarihi")
                            {
                                irsaliyeTarihiSutunu = j;
                            }

                            else if (baslik == "Bakiye")
                            {
                                bakiyeSutunu = j;
                            }
                            else if (baslik == "Sipariş No")
                            {
                                sipNo = j;
                            }
                        }
                    }

                    // "Referans" satırının altındaki verileri DataGridView'e aktar
                    if (baslangicSatiri != -1 && baslangicSutunu != 1 && irsaliyeTarihiSutunu != -1 && bakiyeSutunu != -1 && sipNo != -1)
                    {
                        dataGridView1.Rows.Clear();

                        for (int i = baslangicSatiri; i <= satirSayisi; i++)
                        {
                            string mlzKod = xlRange.Cells[i, baslangicSutunu].Value2?.ToString();
                            string bakiye = xlRange.Cells[i, bakiyeSutunu].Value2?.ToString(); // "Bakiye" sütunundan veri oku
                            string irsaliyeTarihi = xlRange.Cells[i, irsaliyeTarihiSutunu].Value2?.ToString(); // "İrsaliye Tarihi" sütunundan veri oku
                            string spNo = xlRange.Cells[i, sipNo].Value2?.ToString(); // "İrsaliye Tarihi" sütunundan veri oku

                            if (!string.IsNullOrEmpty(mlzKod) && !string.IsNullOrEmpty(bakiye) && !string.IsNullOrEmpty(irsaliyeTarihi))
                            {
                                dataGridView1.Rows.Add(mlzKod, bakiye, irsaliyeTarihi, spNo);
                            }
                            else
                            {
                                // "Mlzkod", "Bakiye", "İrsaliye Tarihi" veya "Miktar" sütunları boşsa, veri okumayı durdur
                                break;
                            }
                        }
                        MessageBox.Show("Aktarım Tamamlandı.");
                    }
                    else
                    {
                        MessageBox.Show("Referans, İrsaliye Tarihi, Miktar veya Bakiye sütunları bulunamadı.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void varmi()
        {
            string sutunAdi = "RefNo";
            // Sütun adının DataGridView içinde zaten var olup olmadığını kontrol et
            bool sütunVarMi = dataGridView1.Columns.Cast<DataGridViewColumn>().Any(col => col.HeaderText == sutunAdi);

            // Eğer sütun yoksa, yeni sütunu ekle
            if (!sütunVarMi)
            {
                DataGridViewTextBoxColumn yeniTextBoxColumn = new DataGridViewTextBoxColumn();
                yeniTextBoxColumn.HeaderText = sutunAdi;
                dataGridView1.Columns.Add(yeniTextBoxColumn);
            }

            string sutunAdix = "RefNo";
            // Sütun adının DataGridView içinde zaten var olup olmadığını kontrol et
            bool sütunVarMix = dataGridView2.Columns.Cast<DataGridViewColumn>().Any(col => col.HeaderText == sutunAdix);

            // Eğer sütun yoksa, yeni sütunu ekle
            if (!sütunVarMix)
            {
                DataGridViewTextBoxColumn yeniTextBoxColumnx = new DataGridViewTextBoxColumn();
                yeniTextBoxColumnx.HeaderText = sutunAdix;
                dataGridView2.Columns.Add(yeniTextBoxColumnx);
            }
        }
        private void varsasil()
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            string silinecekSutunAdi = "RefNo"; // Silinecek sütunun adı
            // Sütun adının DataGridView içinde zaten var olup olmadığını kontrol et
            DataGridViewColumn silinecekSutun = dataGridView1.Columns
                .Cast<DataGridViewColumn>()
                .FirstOrDefault(col => col.HeaderText == silinecekSutunAdi);
            // Eğer sütun varsa, sütunu sil
            if (silinecekSutun != null)
            {
                dataGridView1.Columns.Remove(silinecekSutun);
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            fompak();
        }
        private void doga()
        {
            varmi();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    if (xlRange != null)
                    {
                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;
                        int mlzKodIndex = -1, istenenTarihIndex = -1, miktarIndex = -1, sipNoIndex = -1;

                        for (int j = 1; j <= colCount; j++)
                        {
                            if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null)
                            {
                                if (xlRange.Cells[1, j].Value2.ToString() == "Stok kodu")
                                    mlzKodIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Teslim Tarihi")
                                    istenenTarihIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Kalan miktar")
                                    miktarIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Sipariş No")
                                    sipNoIndex = j;
                            }
                        }

                        if (mlzKodIndex == -1 || istenenTarihIndex == -1 || miktarIndex == -1 || sipNoIndex == -1)
                        {
                            MessageBox.Show("Gerekli başlıklar bulunamadı.");
                            xlWorkbook.Close();
                            xlApp.Quit();
                            return;
                        }
                        for (int i = 2; i <= rowCount; i++)
                        {
                            if (xlRange.Cells[i, mlzKodIndex] != null && xlRange.Cells[i, mlzKodIndex].Value2 != null &&
                                xlRange.Cells[i, miktarIndex] != null && xlRange.Cells[i, miktarIndex].Value2 != null &&
                                xlRange.Cells[i, istenenTarihIndex] != null && xlRange.Cells[i, istenenTarihIndex].Value2 != null &&
                                xlRange.Cells[i, sipNoIndex] != null && xlRange.Cells[i, sipNoIndex].Value2 != null)
                            {
                                string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                                string spNo = xlRange.Cells[i, sipNoIndex].Value2.ToString();
                                double serialDate = Convert.ToDouble(xlRange.Cells[i, istenenTarihIndex].Value2);
                                DateTime istenenTarih = ConvertFromSerialDate(serialDate);
                                string formattedDate = istenenTarih.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                                bool degisti = false;
                                // Excel hücresinin değerini doğrudan okuyarak kontrol et
                                object miktarValue = xlRange.Cells[i, miktarIndex].Value;

                                if (miktarValue != null)
                                {

                                    string miktarString = miktarValue.ToString().Trim(); // Boşlukları temizle
                                    if (double.TryParse(miktarString, out double miktarNumeric))
                                    {

                                        if (mlzKod == "7231537600000")
                                        {
                                            mlzKod = "0723153760000".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "7235366300011") //*
                                        {
                                            mlzKod = "7235366300011".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "7237586000000")
                                        {
                                            mlzKod = "07237586000".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "7489666000000")
                                        {
                                            mlzKod = "7489666/0".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "9010004100000".ToString())
                                        {
                                            mlzKod = "090100041".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "9011196400000")
                                        {
                                            mlzKod = "090111964".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "9521523300000")
                                        {
                                            mlzKod = "095215233".ToString();
                                            degisti = true;
                                        }
                                        else if (mlzKod == "8322550200000")
                                        {
                                            mlzKod = "8322550.2".ToString();
                                            degisti = true;
                                        }
                                        if (degisti == false)
                                        {
                                            dataGridView1.Rows.Add("0" + mlzKod, miktarNumeric, formattedDate, spNo);

                                        }
                                        else
                                        {
                                            dataGridView1.Rows.Add(mlzKod, miktarNumeric, formattedDate, spNo);
                                        }

                                    }
                                    else
                                    {
                                        MessageBox.Show("Miktar değeri geçerli bir sayı değil. Hücre İçeriği: " + miktarString);
                                    }
                                }
                            }
                        }
                        MessageBox.Show("Aktarım Tamamlandı.");
                    }

                    if (xlRange != null)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);
                        xlWorkbook.Close(false, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(xlWorkbook);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }
        private void dog_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            doga();

        }


        private void sumiriko()
        {
            Dictionary<string, double> malzemeKodlari = new Dictionary<string, double>
{
    { "M829SL", 1000 },
    { "N101243", 2500 },
    { "N101562", 2000 },
    { "PH32R0", 2000 },
    { "PX95U0", 2000 },
    { "PY25I0", 2000 },
    { "PY25K0", 5000 },
    { "SK69E0", 3500 },
    { "SK69F0", 5000 },
    { "SK87T0", 1000 },
    { "SL84R0", 1000 },
    { "SL96K0", 1500 },
    { "SM07T0", 2000 },
    { "SM07U0", 750 },
    { "SM07V0", 1750 },
    { "SM15N0", 3000 },
    { "W701890S", 1000 },
    { "W716064S439", 1500 }
            };

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    if (xlRange != null)
                    {
                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;
                        int mlzKodIndex = -1, istenenTarihIndex = -1, miktarIndex = -1, orderTypeIndex = -1;

                        for (int j = 1; j <= colCount; j++)
                        {

                            if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null)
                            {
                                if (xlRange.Cells[1, j].Value2.ToString() == "Item")
                                    mlzKodIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Delivery Date")
                                    istenenTarihIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Ordered quantity")
                                    miktarIndex = j;
                                else if (xlRange.Cells[1, j].Value2.ToString() == "Order type")
                                    orderTypeIndex = j;
                            }
                        }

                        if (mlzKodIndex == -1 || istenenTarihIndex == -1 || miktarIndex == -1 || orderTypeIndex == -1)
                        {
                            MessageBox.Show("Gerekli başlıklar bulunamadı.");
                            xlWorkbook.Close();
                            xlApp.Quit();
                            return;
                        }

                        for (int i = 2; i <= rowCount; i++)
                        {
                            if (xlRange.Cells[i, mlzKodIndex] != null && xlRange.Cells[i, mlzKodIndex].Value2 != null &&
                                xlRange.Cells[i, miktarIndex] != null && xlRange.Cells[i, miktarIndex].Value2 != null &&
                                xlRange.Cells[i, istenenTarihIndex] != null && xlRange.Cells[i, istenenTarihIndex].Value2 != null &&
                                xlRange.Cells[i, orderTypeIndex] != null && xlRange.Cells[i, orderTypeIndex].Value2 != null)
                            {
                                string typIn = xlRange.Cells[i, orderTypeIndex].Value2.ToString();
                                if (typIn == "Executive")
                                {
                                    string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                                    // Tarih dönüşümü buraya ekleniyor

                                    DateTime istenenTarih;
                                    // Kullanıcıdan tarih bilgisini alın
                                    if (malzemeKodlari.TryGetValue(mlzKod, out double malzemeDegeri))
                                    {

                                        // Kullanıcıdan tarih bilgisini alın
                                        if (DateTime.TryParse(xlRange.Cells[i, istenenTarihIndex].Value2.ToString(), out istenenTarih))
                                        {
                                            string formattedDate = istenenTarih.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

                                            // Excel hücresinin değerini doğrudan okuyarak kontrol et
                                            object miktarValue = xlRange.Cells[i, miktarIndex].Value;
                                            double miktarDouble;

                                            // miktarValue'yi double türüne dönüştürme
                                            if (miktarValue != null && double.TryParse(miktarValue.ToString(), out miktarDouble))
                                            {
                                                // İlgili malzeme kodu için işlemi gerçekleştir
                                                double sonuc = miktarDouble / malzemeDegeri;

                                                // Eğer sonucun ondalık kısmı sıfır değilse ve tam bölünmüyorsa
                                                if (sonuc % 1 != 0)
                                                {
                                                    // sonuc'u bir üst tam sayıya yuvarla ve miktarDouble'ı güncelle
                                                    miktarDouble = (int)Math.Ceiling(sonuc) * (int)malzemeDegeri;
                                                    dataGridView1.Rows.Add(mlzKod, miktarDouble, formattedDate);
                                                }
                                                // Eğer sonucun ondalık kısmı sıfır ise ve tam bölünüyorsa
                                                else if (sonuc % 1 == 0)
                                                {
                                                    // miktarDouble'ı güncelle, onu malzemeDegeri'nin katına yuvarla
                                                    miktarDouble = (int)sonuc * (int)malzemeDegeri;
                                                    dataGridView1.Rows.Add(mlzKod, miktarDouble, formattedDate);
                                                }
                                            }
                                        }
                                    }

                                }
                                else if (typIn == "Previsional")
                                {
                                    string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                                    // Tarih dönüşümü buraya ekleniyor

                                    DateTime istenenTarih;
                                    // Kullanıcıdan tarih bilgisini alın
                                    if (malzemeKodlari.TryGetValue(mlzKod, out double malzemeDegeri))
                                    {

                                        // Kullanıcıdan tarih bilgisini alın
                                        if (DateTime.TryParse(xlRange.Cells[i, istenenTarihIndex].Value2.ToString(), out istenenTarih))
                                        {
                                            string formattedDate = istenenTarih.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

                                            // Excel hücresinin değerini doğrudan okuyarak kontrol et
                                            object miktarValue = xlRange.Cells[i, miktarIndex].Value;
                                            double miktarDouble;

                                            // miktarValue'yi double türüne dönüştürme
                                            if (miktarValue != null && double.TryParse(miktarValue.ToString(), out miktarDouble))
                                            {
                                                // İlgili malzeme kodu için işlemi gerçekleştir
                                                double sonuc = miktarDouble / malzemeDegeri;
                                                // Eğer sonucun ondalık kısmı sıfır değilse ve tam bölünmüyorsa
                                                if (sonuc % 1 != 0)
                                                {
                                                    // sonuc'u bir üst tam sayıya yuvarla ve miktarDouble'ı güncelle
                                                    miktarDouble = (int)Math.Ceiling(sonuc) * (int)malzemeDegeri;
                                                    dataGridView2.Rows.Add(mlzKod, miktarDouble, formattedDate);
                                                }
                                                // Eğer sonucun ondalık kısmı sıfır ise ve tam bölünüyorsa
                                                else if (sonuc % 1 == 0)
                                                {
                                                    // miktarDouble'ı güncelle, onu malzemeDegeri'nin katına yuvarla
                                                    miktarDouble = (int)sonuc * (int)malzemeDegeri;
                                                    dataGridView2.Rows.Add(mlzKod, miktarDouble, formattedDate);
                                                }
                                            }
                                        }
                                    }
                                }

                            }
                        }
                        MessageBox.Show("Aktarım Tamamlandı.");
                        if (xlRange != null)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            Marshal.ReleaseComObject(xlRange);
                            Marshal.ReleaseComObject(xlWorksheet);
                            xlWorkbook.Close(false, Type.Missing, Type.Missing);
                            Marshal.ReleaseComObject(xlWorkbook);
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlApp);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }
        private void sumi_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            sumiriko();
        }
        private void cavo()
        {
            varsasil();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int firma = 0;
                int bakiye = 0;
                int kontrol = 0;

                foreach (Excel.Range row in excelRange.Rows)
                {
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Tedarikçi Ürün No")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    for (int jx = 1; jx <= excelRange.Columns.Count; jx++)
                    {
                        string cellValueg = excelRange.Cells[1, jx].Value?.ToString();
                        if (cellValueg != null && cellValueg == "Bakiye")
                        {
                            bakiye = jx + 1;

                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }

                    int startingColumn = bakiye; // Bakiye hücresinin sağında(+1) tarih var ordan itbiaren al
                    for (int j = startingColumn; j <= excelRange.Columns.Count; j++)
                    {
                        for (int i = 1; i <= excelRange.Rows.Count; i++)
                        {
                            object cellValuex = excelRange.Cells[i, j].Value;
                            // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                            if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue))
                            {

                                // Numeric değeri kullanabilirsiniz
                                if (numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();
                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[i, 3].Value.ToString();
                                        
                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        if (firmaDatax== "R0YPS005AF")
                                        {

                                            firmaDatax = "R0YPS005AJ";

                                        }
                                        else if (firmaDatax== " R0YPS008AA")
                                        {
                                            firmaDatax = "R0YPS008";
                                        }
                                        else if ( firmaDatax == "")
                                        {
                                            firmaDatax = "R0YPS008AB";
                                        }
                                       
                                        // DataGridView'e ekle, bu sefer format dönüşümü yaparak ekleyin
                                        dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                    }
                                }
                            }
                        }
                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    break;
                }
                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void cav_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            cavo();
        }
        private void odelo()
        {
            varsasil();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int firma = 0;
                int bakiye = 0;
                int kontrol = 0;

                foreach (Excel.Range row in excelRange.Rows)
                {
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Parça No")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    for (int jx = 1; jx <= excelRange.Columns.Count; jx++)
                    {
                        string cellValueg = excelRange.Cells[1, jx].Value?.ToString();
                        if (cellValueg != null && cellValueg == "Bakiye")
                        {
                            bakiye = jx + 1;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }

                    int startingColumn = bakiye; // Bakiye hücresinin sağında(+1) tarih var, oradan itibaren al
                    for (int j = startingColumn; j <= excelRange.Columns.Count; j++)
                    {
                        for (int i = 2; i <= excelRange.Rows.Count; i++)
                        {
                            object cellValuex = excelRange.Cells[i, j].Value;
                            // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa

                            if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue))
                            {
                                // Numeric değeri kullanabilirsiniz
                                if (numericValue != 0)
                                {
                                    // Sağındaki hücrede "X" kontrolü yap
                                    string rightCellValue = excelRange.Cells[i, j + 1].Value?.ToString();
                                    if (rightCellValue != null && rightCellValue.Trim().ToUpper() == "X")
                                    {
                                        // "X" varsa numericValue'yi yaz
                                        string dateValue = excelRange.Cells[1, j].Value?.ToString();
                                        if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                        {
                                            string firmaDatax = excelRange.Cells[i, 2].Value.ToString();
                                            string formattedDate = date.ToString("dd.MM.yyyy");
                                            dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                    else   // ÖNGÖRÜ OLAYI
                                    {
                                        // "X" yoksa numericValue'yi  ekleyerek yaz
                                        string dateValue = excelRange.Cells[1, j].Value?.ToString();
                                        if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                        {
                                            string firmaDatax = excelRange.Cells[i, 2].Value.ToString();
                                            string formattedDate = date.ToString("dd.MM.yyyy");
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    break;
                }
                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }

        }

        private void odel_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            odelo();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        private void rollmech()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 6;

                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Ürün No") 
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }
                    Dictionary<string, double> yesilMiktarlar = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();

                    int startingColumn = bakiye;
                    
                    // Yeşil renkli hücre işlemleri // ÖNGÖRÜ İŞLEMLERİ
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 147, 255, 149)))
                        {
                            for (int j = cell.Column; j <= excelRange.Columns.Count; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;

                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        // Veri daha önce eklenmemişse ekle ve geçici listeyi güncelle
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                            if (firmaDatax == "3A00 217 006-01")
                                            {
                                                firmaDatax = "3A00 217 006-01";
                                            }
                                            else if (firmaDatax == "3A00 310 037-01")
                                            {
                                                firmaDatax = "3A00 310 037";
                                            }
                                            else if (firmaDatax == "3A00 217 008-01")
                                            {
                                                firmaDatax = "3A00 217 008";
                                            }
                                            else if (firmaDatax == "3900 310 009-01")
                                            {
                                                firmaDatax = "3900 310 009";
                                            }
                                            else if (firmaDatax == "R0YPS008AA")
                                            {
                                                firmaDatax = "R0YPS008";
                                            }
                                            // DataGridView'e ekle, bu sefer format dönüşümü yaparak ekleyin //ÖNGÖRÜ
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, formattedDate);
                                           
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Diğer renkli hücre işlemleri
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 180, 214, 238)) ||
                            cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 174, 137)))
                        {

                            for (int j = startingColumn; j <= excelRange.Columns.Count - 1; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;
                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                        
                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");

                                        // Veri daha önce eklenmemişse ekle ve geçici listeyi güncelle
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                            if (firmaDatax == "3A00 217 006-01")
                                            {
                                                firmaDatax = "3A00 217 006-01";
                                            }
                                            else if (firmaDatax == "3A00 310 037-01")
                                            {
                                                firmaDatax = "3A00 310 037";
                                            }
                                            else if (firmaDatax == "3A00 217 008-01")
                                            {
                                                firmaDatax = "3A00 217 008";
                                            }
                                            else if (firmaDatax == "3900 310 009-01")
                                            {
                                                firmaDatax = "3900 310 009";
                                            }
                                            else if (firmaDatax == "R0YPS008AA")
                                            {
                                                firmaDatax = "R0YPS008";
                                            }

                                            dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("Aktarım Tamamlandı.");
                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            rollmech();
        }
        private void farba()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 0;
                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Malzeme No") // Değiştirmeniz gerekebilir
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "SAS Ölçü Br.") // Değiştirmeniz gerekebilir
                        {
                            bakiye = j + 1;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }
                    Dictionary<string, double> yesilMiktarlar = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();
                    HashSet<string> eklenenVeriler2 = new HashSet<string>();
                    int startingColumn = bakiye; // Bakiye hücresinin sağında(+1) tarih var ordan itibaren al

                    // Diğer renkli hücre işlemleri
                    foreach (Excel.Range cell in row.Cells)
                    {
                        // Sadece yeşil renkteki hücreyi kontrol et
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 255, 0)))
                        {
                            // Diğer işlemleri gerçekleştir...
                            int j = cell.Column; // Yeşil hücrenin sütun numarasını al
                            object cellValuex = excelRange.Cells[row.Row, j].Value;
                            if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                            {
                                // Diğer işlemleri gerçekleştir...
                                string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                {
                                    string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                    string formattedDate = date.ToString("dd.MM.yyyy");
                                    string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";

                                    if (!eklenenVeriler.Contains(uniqueKey))
                                    {
                                        eklenenVeriler.Add(uniqueKey);
                                        dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                    }
                                }
                            }
                        }
                        else if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255)))
                        {
                            int j2 = cell.Column; // beyaz hücrenin sütun numarasını al
                            object cellValuex2 = excelRange.Cells[row.Row, j2].Value;
                            if (cellValuex2 != null && double.TryParse(cellValuex2.ToString(), out double numericValue2) && numericValue2 != 0)
                            {
                                // Diğer işlemleri gerçekleştir...
                                string dateValue2 = excelRange.Cells[1, j2].Value?.ToString();

                                if (!string.IsNullOrEmpty(dateValue2) && DateTime.TryParse(dateValue2, out DateTime date2))
                                {
                                    string firmaDatax2 = excelRange.Cells[row.Row, firma].Value.ToString();
                                    string formattedDate2 = date2.ToString("dd.MM.yyyy");
                                    string uniqueKey2 = $"{firmaDatax2}_{numericValue2}_{formattedDate2}";

                                    if (!eklenenVeriler2.Contains(uniqueKey2))
                                    {
                                        eklenenVeriler2.Add(uniqueKey2);
                                        dataGridView2.Rows.Add(firmaDatax2, numericValue2, formattedDate2);
                                    }
                                }
                            }
                        }
                        
                    }
                }
                MessageBox.Show("Aktarım Tamamlandı.");

                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void button4_Click_2(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            farba();
        }
        private void pressan_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            verigetir();
        }

        private void orhanOdelo()
        {
            varsasil();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int firma = 0;
                int bakiye = 0;
                int kontrol = 0;

                foreach (Excel.Range row in excelRange.Rows)
                {
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[7, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Parça No")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    int gercek = 0;
                    int intValue = 0;
                    for (int jx = 1; jx <= excelRange.Columns.Count; jx++)
                    {
                        string cellValueg = excelRange.Cells[7, jx].Value?.ToString();
                        if (cellValueg != null && cellValueg == "Bakiye")
                        {
                            bakiye = jx + 1;
                            kontrol = 1;
                            gercek = jx;
                            break;
                        }
                        
                        kontrol++;
                    }
                   
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }

                    int startingColumn = bakiye;
                    double numericValue=0;
                    // Bakiye hücresinin sağında(+1) tarih var, oradan itibaren al
                    for (int j = startingColumn; j <= excelRange.Columns.Count; j++)
                    {
                        for (int i = 1; i <= excelRange.Rows.Count; i++)
                        {
                          

                          
                            object cellValuex = excelRange.Cells[i, j].Value;

                            // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                            if (cellValuex != null && double.TryParse(cellValuex.ToString(), out numericValue))
                            {

                                // Numeric değeri kullanabilirsiniz
                                if (numericValue != 0)
                                {
                                    object cellValue = excelRange.Cells[i, gercek].Value2;
                                    if (cellValue != null)
                                    {

                                        if (cellValue is double)
                                        {
                                            double doubleValue = (double)cellValue;
                                            intValue = (int)doubleValue;

                                            // intValue'i kullanabilirsiniz
                                        }
                                        else
                                        {
                                            // Hücrede sayısal bir değer yoksa veya double'a dönüştürülemezse bir işlem yapabilirsiniz.
                                        }
                                    }
                                    DateTime today = DateTime.Today;
                                    DateTime firstSaturday = today;

                                    // Bugün Cumartesi değilse, bir sonraki Cumartesi'yi bul
                                    if (today.DayOfWeek != DayOfWeek.Saturday)
                                    {
                                        int daysUntilNextSaturday = ((int)DayOfWeek.Saturday - (int)today.DayOfWeek + 7) % 7;
                                        firstSaturday = today.AddDays(daysUntilNextSaturday);
                                    }
                                    // Sağındaki hücrede "X" kontrolü yap
                                    string rightCellValue = excelRange.Cells[i, j + 1].Value?.ToString();
                                    if (rightCellValue != null && rightCellValue.Trim().ToUpper() == "X")
                                    {
                                        numericValue += intValue;

                                        // "X" varsa numericValue'yi yaz
                                        string dateValue = excelRange.Cells[7, j].Value?.ToString();
                                        if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                        {
                                            string firmaDatax = excelRange.Cells[i, 2].Value.ToString();
                                            string formattedDate = date.ToString("dd.MM.yyyy");
                                            dataGridView1.Rows.Add(firmaDatax, numericValue, firstSaturday.ToString("dd.MM.yyyy"));
                                        }
                                    }
                                    else  // ÖNGÖRÜ OLAYI
                                    {
                                        
                                        // "X" yoksa numericValue'yi  ekleyerek yaz
                                        string dateValue = excelRange.Cells[7, j].Value?.ToString();
                                        if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                        {
                                            string firmaDatax = excelRange.Cells[i, 2].Value.ToString();
                                            string formattedDate = date.ToString("dd.MM.yyyy");
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    break;
                }
                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void orhnOdel_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();
            orhanOdelo();
        }
        private void akaoglu(){
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Excel Dosyası Seç";
                openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog1.RestoreDirectory = true;
                DateTime today = DateTime.Now;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int mlzKodIndex = -1;

                    // 2. satırdan "ÜRÜN KODU" başlığını bul
                    for (int col = 1; col <= xlRange.Columns.Count; col++)
                    {
                        if (xlRange.Cells[3, col].Value2 != null && xlRange.Cells[3, col].Value2.ToString() == "Stok Kod")
                        {
                            mlzKodIndex = col;
                            break;
                        }
                    }

                    if (mlzKodIndex == -1)
                    {
                        MessageBox.Show("Gerekli başlık bulunamadı.");
                        xlWorkbook.Close();
                        xlApp.Quit();
                        return;
                    }


                    int rowCount = xlRange.Rows.Count;
                    string kodx = null;
                    DateTime ayBilgi = DateTime.MinValue;
                    string formatliTarih = "";
                    for (int i = 1; i <= rowCount; i++) // 2. satırı atladık
                    {
                        if (xlRange.Cells[i, mlzKodIndex] != null && xlRange.Cells[i, mlzKodIndex].Value2 != null)
                        {
                            string mlzKod = xlRange.Cells[i, mlzKodIndex].Value2.ToString();
                            for (int j = mlzKodIndex + 1; j <= xlRange.Columns.Count; j++)
                            {
                                if (xlRange.Cells[2, j] != null && xlRange.Cells[2, j].Value2 != null)
                                {
                                    string tarihStr = xlRange.Cells[2, j].Text;
                                    string veri = xlRange.Cells[i, j].Value2 != null ? xlRange.Cells[i, j].Value2.ToString() : "";
                                    DateTime tarih;
                                    if (DateTime.TryParse(tarihStr, out tarih))
                                    {
                                        formatliTarih = tarih.ToString("dd.MM.yyyy"); // Tarihi istediğiniz formatta stringe çevir
                                    }
                                    else
                                    {
                                        MessageBox.Show("Geçerli bir tarih değil.");
                                    }


                                    // Eğer Excel sayfasındaki başlık içerisinde ay ifadesi varsa, DataGridView'e ekle

                                    kodx = veri;
                                    //    MessageBox.Show(kodx);

                                    if (int.TryParse(kodx, out int mlzKodInt) && mlzKodInt != 0)
                                    {
                                        if (formatliTarih != null)
                                        {
                                            if (mlzKod == "257150T010")
                                            {
                                                mlzKod = "25715-0T010".ToString();
                                            }
                                            dataGridView1.Rows.Add(mlzKod, veri, formatliTarih);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {

                        }
                    }
                    MessageBox.Show("Aktarım Tamamlandı.");
                    if (xlRange != null)
                    {
                        Marshal.ReleaseComObject(xlRange);
                    }

                    if (xlWorksheet != null)
                    {
                        Marshal.ReleaseComObject(xlWorksheet);
                    }

                    xlWorkbook.Close(false, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(xlWorkbook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }

        }
        private void button5_Click(object sender, EventArgs e)
        {
            akaoglu();
        }
      
        private void cavoRenkli()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 5;

                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Ürün No")
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }
                    Dictionary<string, double> yesilMiktarlar = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();

                    int startingColumn = bakiye;

                    // Yeşil renkli hücre işlemleri // ÖNGÖRÜ İŞLEMLERİ
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 238, 198, 132)))
                        {
                            for (int j = cell.Column; j <= excelRange.Columns.Count; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;

                                // Eğer hücre değeri null değilse ve bir sayısal değeri temsil ediyorsa
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    // En üst sütundaki tarih değerini al
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                    // Eğer tarih değeri null değilse ve dd.MM.yyyy HH:mm:ss formatına uyuyorsa
                                    if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                    {
                                        // İlk sütundaki firma verisini al
                                        string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                        // Tarihi istediğiniz formata çevirin
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        // Veri daha önce eklenmemişse ekle ve geçici listeyi güncelle
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                          
                                            // DataGridView'e ekle, bu sefer format dönüşümü yaparak ekleyin //ÖNGÖRÜ
                                            dataGridView2.Rows.Add(firmaDatax, numericValue, formattedDate);

                                        }
                                    }
                                }
                            }
                        }
                    }
                    foreach (Excel.Range cell in row.Cells)
                    {
                        if ((cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 180, 214, 238))) ||
                            (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 174, 137))))
                        {
                            for (int j = startingColumn; j <= excelRange.Columns.Count - 1; j++)
                            {
                                object cellValuex = excelRange.Cells[row.Row, j].Value;
                                if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                                {
                                    string dateValue = excelRange.Cells[1, j].Value?.ToString();
                                    if (!string.IsNullOrEmpty(dateValue) && (DateTime.TryParse(dateValue, out DateTime date) || dateValue == "Bakiye"))
                                    {
                                        string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                        string formattedDate = date.ToString("dd.MM.yyyy");
                                        if (dateValue == "Bakiye")
                                        {
                                            double excelDateValue;
                                            Excel.Range g1range = excelWorksheet.Range["G1"];
                                            string g1r = g1range.Value2.ToString();
                                            if (double.TryParse(g1r, out excelDateValue))
                                            {
                                                DateTime g1Date = DateTime.FromOADate(excelDateValue);
                                                formattedDate = g1Date.ToString("dd.MM.yyyy");
                                            }
                                        }
                                        string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";
                                        if (!eklenenVeriler.Contains(uniqueKey))
                                        {
                                            eklenenVeriler.Add(uniqueKey);
                                            dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("Aktarım Tamamlandı.");
                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void button6_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            cavoRenkli();
        }
        private void aygersan()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Excel Dosyası Seç";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(selectedFileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                List<string> mlzKodList = new List<string>();
                List<string> istenenTarihList = new List<string>();
                List<string> miktarList = new List<string>();
                int kontrol = 0;
                int bakiye = 0;
                foreach (Excel.Range row in excelRange.Rows)
                {
                    int firma = 0;
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "Malzeme No") // Değiştirmeniz gerekebilir
                        {
                            firma = j;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        string cellValue = excelRange.Cells[1, j].Value?.ToString();
                        if (cellValue != null && cellValue == "SAS Ölçü Br.") // Değiştirmeniz gerekebilir
                        {
                            bakiye = j + 1;
                            kontrol = 1;
                            break;
                        }
                        kontrol++;
                    }
                    if (kontrol != 1)
                    {
                        MessageBox.Show("Yanlış Dosya Seçildi.");
                        break;
                    }
                    Dictionary<string, double> yesilMiktarlar = new Dictionary<string, double>();
                    string ustSatirBilgi = string.Empty;
                    HashSet<string> eklenenVeriler = new HashSet<string>();
                    HashSet<string> eklenenVeriler2 = new HashSet<string>();
                    int startingColumn = bakiye; // Bakiye hücresinin sağında(+1) tarih var ordan itibaren al

                    // Diğer renkli hücre işlemleri
                    foreach (Excel.Range cell in row.Cells)
                    {
                        // Sadece yeşil renkteki hücreyi kontrol et
                        if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 255, 0)))
                        {
                            // Diğer işlemleri gerçekleştir...
                            int j = cell.Column; // Yeşil hücrenin sütun numarasını al
                            object cellValuex = excelRange.Cells[row.Row, j].Value;
                            if (cellValuex != null && double.TryParse(cellValuex.ToString(), out double numericValue) && numericValue != 0)
                            {
                                // Diğer işlemleri gerçekleştir...
                                string dateValue = excelRange.Cells[1, j].Value?.ToString();

                                if (!string.IsNullOrEmpty(dateValue) && DateTime.TryParse(dateValue, out DateTime date))
                                {
                                    string firmaDatax = excelRange.Cells[row.Row, firma].Value.ToString();
                                    string formattedDate = date.ToString("dd.MM.yyyy");
                                    string uniqueKey = $"{firmaDatax}_{numericValue}_{formattedDate}";

                                    if (!eklenenVeriler.Contains(uniqueKey))
                                    {
                                        eklenenVeriler.Add(uniqueKey);
                                        dataGridView1.Rows.Add(firmaDatax, numericValue, formattedDate);
                                    }
                                }
                            }
                        }
                        else if (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255))||
                                (cell.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(245, 245, 245))))
                        {
                            int j2 = cell.Column; // beyaz hücrenin sütun numarasını al
                            object cellValuex2 = excelRange.Cells[row.Row, j2].Value;
                            if (cellValuex2 != null && double.TryParse(cellValuex2.ToString(), out double numericValue2) && numericValue2 != 0)
                            {
                                // Diğer işlemleri gerçekleştir...
                                string dateValue2 = excelRange.Cells[1, j2].Value?.ToString();

                                if (!string.IsNullOrEmpty(dateValue2) && DateTime.TryParse(dateValue2, out DateTime date2))
                                {
                                    string firmaDatax2 = excelRange.Cells[row.Row, firma].Value.ToString();
                                    string formattedDate2 = date2.ToString("dd.MM.yyyy");
                                    string uniqueKey2 = $"{firmaDatax2}_{numericValue2}_{formattedDate2}";

                                    if (!eklenenVeriler2.Contains(uniqueKey2))
                                    {
                                        eklenenVeriler2.Add(uniqueKey2);
                                        dataGridView2.Rows.Add(firmaDatax2, numericValue2, formattedDate2);
                                    }
                                }
                            }
                        }

                    }
                }
                MessageBox.Show("Aktarım Tamamlandı.");

                excelWorkbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi.");
            }
        }
        private void Aygers_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            aygersan();
        }


        private void opsanYni()
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            varsasil();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;

                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel.Range xlRange = null;

                try
                {
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(filePath);
                    xlWorksheet = xlWorkbook.Sheets[1];

                    // İlk dolu satırın indeksini bul
                    int firstDataRowIndex = FindFirstDataRowIndex(xlWorksheet);

                    if (firstDataRowIndex > 0)
                    {
                        // Var olan DataGridView kontrolüne ekle
                        for (int i = firstDataRowIndex; i <= xlWorksheet.UsedRange.Rows.Count; i++)
                        {
                            xlRange = xlWorksheet.Rows[i];

                            // İlk hücrenin değerini "MlzKod" sütununa ekle
                            string mlzKodValue = xlRange.Cells[1, 1].Text;

                            // Diğer hücreleri "Miktar" sütununa ekle
                            for (int j = 2; j <= xlWorksheet.UsedRange.Columns.Count; j++)
                            {
                                string miktarValue = xlRange.Cells[1, j].Text;
                                // İlgili tarih sütununu belirle
                                string taih = xlWorksheet.Cells[4, j].Text;
                                // Noktaları temizle
                                miktarValue = miktarValue.Replace(".", "");

                                if (mlzKodValue == "113-01-70055724")
                                {
                                    mlzKodValue = "70055724";
                                }
                                else if (mlzKodValue == "113-01-PZ31-K624B93AP10A")
                                {
                                    mlzKodValue = "PZ31-K624B93-A-PIA-10-10";
                                }
                                else if (mlzKodValue == "113-06-W720476")
                                {
                                    mlzKodValue = "W720476";
                                }
                                else if (mlzKodValue == "113-01-H1BB-109A26APIA11")
                                {
                                    mlzKodValue = "H1BB-109A26-APIA-11";
                                }
                                else if (mlzKodValue == "113-05-W713648-S")
                                {
                                    mlzKodValue = "W713648";
                                }
                                if (miktarValue != "" && taih != "" && taih !="MOQ")
                                {
                                    if (DateTime.TryParseExact(taih, taih.StartsWith("0") ? "dd.MM.yyyy" : "d.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))

                                    {
                                        // Tarihten 2 gün çıkar
                                        DateTime yeniTarih = date.AddDays(-2);

                                        // Yeni tarihi istenen formata dönüştür
                                        string yeniTaih = yeniTarih.ToString("dd.MM.yyyy");

                                        // Yeni tarihi dataGridView1'e ekle
                                        if (miktarValue != "" && yeniTaih != "" && yeniTaih != "MOQ")
                                        {
                                            dataGridView1.Rows.Add(mlzKodValue, miktarValue, yeniTaih);
                                        }
                                    }
                                }
                            }
                        }

                        MessageBox.Show("Aktarım Tamamlandı.");
                    }
                    else
                    {
                        MessageBox.Show("Dolu satır bulunamadı.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Excel işlemleri sırasında bir hata oluştu: " + ex.Message);
                }
                if (xlApp != null)
                {
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // Kullanılan Excel nesnelerini serbest bırak
                    ReleaseObject(xlWorksheet);
                    ReleaseObject(xlWorkbook);
                    ReleaseObject(xlApp);

                    // Garbage Collector'ı çağır
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }


        }
        private void opsanY_Click(object sender, EventArgs e)
        {
            opsanYni();
        }
    }
}