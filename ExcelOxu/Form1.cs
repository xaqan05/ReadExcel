using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelOxu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //avromed region


            //string NormalizeMedName(string raw)
            //{
            //    int lastStart = raw.LastIndexOf("(");
            //    int lastEnd = raw.LastIndexOf(")");
            //    if (lastStart >= 0 && lastEnd > lastStart)
            //        raw = raw.Remove(lastStart, lastEnd - lastStart + 1);

            //    if (raw.Contains("№"))
            //    {
            //        raw = raw.Substring(0, raw.IndexOf("№")).Trim();
            //    }

            //    return raw.Trim().ToLowerInvariant();
            //}

            //string NormalizeText(string text)
            //{
            //    return text.ToLowerInvariant()
            //        .Replace("ə", "e")
            //        .Replace("ı", "i")
            //        .Replace("ö", "o")
            //        .Replace("ü", "u")
            //        .Replace("ç", "c")
            //        .Replace("ş", "s")
            //        .Replace("ğ", "g");
            //}

            //using var ofd = new OpenFileDialog()
            //{
            //    Filter = "Excel Dosyaları|*.xlsx;*.xls",
            //    Title = "Excel Dosyası Seçin"
            //};

            //if (ofd.ShowDialog() != DialogResult.OK)
            //    return;

            //using var wb = new XLWorkbook(ofd.FileName);
            //var ws = wb.Worksheet(1);

            //var dt = new DataTable();
            //dt.Columns.Add("Dərman adı", typeof(string));
            //List<string> prevCities = [];
            //Dictionary<string, int> medicines = [];
            //int lastColumn = 0;
            //bool isSpecial = false;
            //DataTable specialMeds = new();
            //specialMeds.Columns.Add("Dərman adı", typeof(string));
            //specialMeds.Columns.Add("Şəhər və ya Filial", typeof(string));
            //specialMeds.Columns.Add("Say", typeof(double));
            //string lastBranchName = "";
            //string lastMedName = "";

            //foreach (var row in ws.RowsUsed())
            //{
            //    int outline = row.OutlineLevel;

            //    for (int i = 1; i <= 5; i++)
            //    {
            //        var val = row.Cell(i).GetString().Trim();
            //        if (!string.IsNullOrEmpty(val) && val.ToLower().Contains("filial"))
            //        {
            //            lastBranchName = val;
            //            break;
            //        }
            //    }

            //    if (string.IsNullOrWhiteSpace(lastBranchName))
            //        lastBranchName = "Bilinməyən Filial";
            //    if (outline == 1)
            //    {

            //        string cityName = row.Cell(2).GetString();

            //        if (cityName.Contains("Şəxsi apteklər"))
            //        {
            //            isSpecial = true;
            //        }
            //        else
            //        {
            //            dt.Columns.Add(cityName, typeof(double));
            //            lastColumn = dt.Columns.Count - 1;
            //            isSpecial = false;
            //        }
            //    }
            //    else if (outline == 2)
            //    {
            //        string medNameRaw = row.Cell(2).GetString().Trim();
            //        if (isSpecial)
            //        {
            //            lastMedName = medNameRaw;
            //            continue;
            //        }

            //        string medName = NormalizeMedName(medNameRaw);
            //        double medCount = row.Cell(3).GetValue<double?>() ?? 0;

            //        if (!medicines.ContainsKey(medName))
            //        {
            //            var dataRow = dt.NewRow();
            //            dataRow[0] = medName;
            //            for (int i = 1; i < dt.Columns.Count; i++)
            //                dataRow[i] = 0;

            //            dataRow[lastColumn] = medCount;
            //            dt.Rows.Add(dataRow);
            //            medicines.Add(medName, dt.Rows.Count - 1);
            //        }
            //        else
            //        {
            //            var dr = dt.Rows[medicines[medName]];
            //            double currentVal = double.TryParse(dr[lastColumn]?.ToString(), out var cur) ? cur : 0;
            //            dr[lastColumn] = currentVal + medCount;
            //        }
            //    }
            //    else if (outline == 3 && isSpecial)
            //    {

            //        string rawMedName = lastMedName;
            //        string rawPharmacyInfo = row.Cell(2).GetString().Trim();

            //        double count = row.Cell(3).GetValue<double?>() ?? 0;

            //        string medName = NormalizeMedName(rawMedName);

            //        string[] possibleCities = new[] {
            //            "Ağcabədi", "Ağdam", "Beyləqan", "Bərdə", "Füzuli", "Tərtər",
            //            "Gədəbəy", "Gəncə", "Goranboy", "Naftalan", "Şəmkir-Çinarlı", "Şəmkir",
            //            "Astara", "Biləsuvar", "Cəlilabad", "Lerik", "Lənkəran", "Masallı", "Yardımlı",
            //            "Ağdaş", "Göyçay", "Mingəçevir", "Ucar", "Yevlax", "Zərdab",
            //            "Alıcılar",
            //            "Ağstafa", "Qazax", "Tovuz",
            //            "Balakən", "İsmayıllı", "Qax", "Qəbələ", "Şəki", "Zaqatala","Oğuz Xaçmaz kəndi",
            //            "Hacıqabul", "İmişli", "Kürdəmir", "Neftçala", "Saatlı", "Sabirabad", "Salyan", "Şamaxı", "Şirvan",
            //            "Quba", "Qusar", "Şabran", "Siyezen", "Xaçmaz", "Xudat"
            //        };

            //        string normPharmacy = NormalizeText(rawPharmacyInfo);
            //        string extractedCity = "";
            //        int bestPos = int.MaxValue;

            //        foreach (var city in possibleCities)
            //        {
            //            string normCity = NormalizeText(city);
            //            int pos = normPharmacy.IndexOf(normCity);
            //            if (pos >= 0 && pos < bestPos)
            //            {
            //                bestPos = pos;
            //                extractedCity = city;
            //            }
            //        }

            //        string location;
            //        if (!string.IsNullOrEmpty(extractedCity))
            //        {
            //            location = extractedCity;
            //        }
            //        else
            //        {
            //            location = $"{lastBranchName.Trim()} - {rawPharmacyInfo.Trim()}";
            //        }

            //        var dr = specialMeds.NewRow();
            //        dr[0] = medName;
            //        dr[1] = location;
            //        dr[2] = count;
            //        specialMeds.Rows.Add(dr);
            //    }
            //}

            //DataTable finalTable = new();
            //finalTable.Columns.Add("Dərman adı", typeof(string));
            //finalTable.Columns.Add("Şəhər və ya Filial", typeof(string));
            //finalTable.Columns.Add("Say", typeof(double));

            //string[] notCities = ["ampula", "məhlul", "gel", "tabletka", "şampun", "kapsul", "aerozol", "damcı", "sprey", "krem"];
            //Dictionary<string, double> totals = new();

            //void AddOrUpdate(string medName, string city, double count)
            //{
            //    if (!string.IsNullOrWhiteSpace(city))
            //        city = char.ToUpper(city[0]) + city[1..].ToLowerInvariant();

            //    if (string.IsNullOrWhiteSpace(city) || notCities.Any(f => city.ToLowerInvariant().Contains(f)))
            //        return;

            //    string key = $"{medName.ToLowerInvariant()}|{city.ToLowerInvariant()}";
            //    if (totals.ContainsKey(key))
            //        totals[key] += count;
            //    else
            //        totals[key] = count;
            //}

            //foreach (DataRow dr in dt.Rows)
            //{
            //    string medName = dr[0].ToString();
            //    for (int i = 1; i < dt.Columns.Count; i++)
            //    {
            //        string city = dt.Columns[i].ColumnName;
            //        double count = double.TryParse(dr[i]?.ToString(), out var c) ? c : 0;
            //        if (count != 0)
            //            AddOrUpdate(medName, city, count);
            //    }
            //}

            //foreach (DataRow dr in specialMeds.Rows)
            //{
            //    string medName = dr[0].ToString();
            //    string city = dr[1].ToString();
            //    double count = double.TryParse(dr[2]?.ToString(), out var c) ? c : 0;
            //    if (count != 0)
            //        AddOrUpdate(medName, city, count);
            //}

            //foreach (var kvp in totals)
            //{
            //    var parts = kvp.Key.Split('|');
            //    var row = finalTable.NewRow();
            //    row["Dərman adı"] = parts[0];
            //    row["Şəhər və ya Filial"] = parts[1];
            //    row["Say"] = kvp.Value;
            //    finalTable.Rows.Add(row);
            //}

            //DataTable pivotTable = new();
            //pivotTable.Columns.Add("Dərman adı", typeof(string));

            //var uniqueLocations = finalTable.AsEnumerable()
            //    .Select(r => r.Field<string>("Şəhər və ya Filial"))
            //    .Distinct()
            //    .OrderBy(c => c)
            //    .ToList();

            //foreach (var location in uniqueLocations)
            //    pivotTable.Columns.Add(location, typeof(double));

            //var uniqueMeds = finalTable.AsEnumerable()
            //    .Select(r => r.Field<string>("Dərman adı"))
            //    .Distinct()
            //    .OrderBy(m => m)
            //    .ToList();

            //foreach (var med in uniqueMeds)
            //{
            //    var newRow = pivotTable.NewRow();
            //    newRow["Dərman adı"] = med;

            //    foreach (var location in uniqueLocations)
            //    {
            //        var match = finalTable.AsEnumerable()
            //            .FirstOrDefault(r => r.Field<string>("Dərman adı") == med && r.Field<string>("Şəhər və ya Filial") == location);

            //        newRow[location] = match != null ? match.Field<double>("Say") : 0;
            //    }

            //    pivotTable.Rows.Add(newRow);
            //}

            //using (var sfd = new SaveFileDialog
            //{
            //    Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
            //    Title = "Nəticəni Excel olaraq Yadda Saxla",
            //    FileName = "AvromedRegionSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            //})
            //{
            //    if (sfd.ShowDialog() == DialogResult.OK)
            //    {
            //        using var saveWb = new XLWorkbook();
            //        var ws1 = saveWb.Worksheets.Add("Pivot_Hesabatı");

            //        for (int c = 0; c < pivotTable.Columns.Count; c++)
            //        {
            //            ws1.Cell(1, c + 1).Value = pivotTable.Columns[c].ColumnName;
            //        }

            //        for (int r = 0; r < pivotTable.Rows.Count; r++)
            //        {
            //            for (int c = 0; c < pivotTable.Columns.Count; c++)
            //            {
            //                var cell = ws1.Cell(r + 2, c + 1);
            //                var value = pivotTable.Rows[r][c];

            //                if (value is double d)
            //                {
            //                    cell.Value = d;
            //                }
            //                else
            //                {
            //                    cell.Value = value?.ToString() ?? "";
            //                }
            //            }
            //        }

            //        ws1.Columns().AdjustToContents();
            //        saveWb.SaveAs(sfd.FileName);

            //        MessageBox.Show(
            //            "Nəticə saxlanıldı:\n" + sfd.FileName,
            //            "Uğurlu",
            //            MessageBoxButtons.OK,
            //            MessageBoxIcon.Information
            //        );
            //    }
            //}

            //dataGridView1.DataSource = pivotTable;
            //dataGridView1.Visible = false;
            //dataGridView1.SendToBack();




            //avromed tam versiya
            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Faylını Seç"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string[] possibleCities = new[]
            {
        "Sumqayıt",
        "Mingəçevir",
        "Lənkəran",
        "Yevlax",
        "Naftalan",
        "Şəki",
        "Şirvan",
        "Xankəndi",
        "Ağcabədi",
        "Ağdam",
        "Ağdaş",
        "Ağstafa",
        "Ağsu",
        "Astara",
        "Balakən",
        "Bərdə",
        "Beyləqan",
        "Biləsuvar",
        "Cəlilabad",
        "Daşkəsən",
        "Füzuli",
        "Gəncə",
        "Gədəbəy",
        "Goranboy",
        "Göyçay",
        "Göygöl",
        "Hacıqabul",
        "Xaçmaz",
        "İmişli",
        "İsmayıllı",
        "Kəlbəcər",
        "Kürdəmir",
        "Qax",
        "Qazax",
        "Qəbələ",
        "Qobustan",
        "Quba",
        "Qubadlı",
        "Qusar",
        "Laçın",
        "Lerik",
        "Masallı",
        "Neftçala",
        "Oğuz",
        "Saatlı",
        "Sabirabad",
        "Salyan",
        "Samux",
        "Siyəzən",
        "Şabran",
        "Şamaxı",
        "Şəmkir",
        "Şuşa",
        "Tərtər",
        "Tovuz",
        "Ucar",
        "Yardımlı",
        "Zaqatala",
        "Zəngilan",
        "Zərdab",
        "Naxçıvan",
        "Xudat"
    };

            Dictionary<string, string> cityMappings = new Dictionary<string, string>();

            try
            {
                string mappingFilePath = Path.Combine(Application.StartupPath, "mappings", "avromed.json");

                if (File.Exists(mappingFilePath))
                {
                    string jsonString = File.ReadAllText(mappingFilePath);
                    cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
                }
                else
                {
                    MessageBox.Show($"Haritalama faylı tapılmadı: {mappingFilePath}. Yalnız əsas şəhərlər siyahısı istifadə ediləcək.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                using var wb = new XLWorkbook(ofd.FileName);
                var ws = wb.Worksheet(1);

                var rows = ws.RowsUsed().Skip(1);
                var salesDataList = new List<SatisVerisi>();

                foreach (var row in rows)
                {
                    var address = row.Cell(7).GetString();
                    string drugNameRaw = row.Cell(10).GetString();

                    string drugNameClean = drugNameRaw.Trim();
                    drugNameClean = Regex.Replace(drugNameClean, @"\s+", " ").Trim();

                    if (row.Cell(14).TryGetValue(out decimal soldQuantity))
                    {
                        string identifiedRegion = null;

                        var mappedCityPair = cityMappings
                            .FirstOrDefault(map => address.Contains(map.Key));

                        if (mappedCityPair.Key != null)
                        {
                            identifiedRegion = mappedCityPair.Value;
                        }

                        if (identifiedRegion == null)
                        {
                            var cityCandidate = possibleCities.FirstOrDefault(c => address.Contains(c));

                            if (cityCandidate != null)
                            {
                                identifiedRegion = cityCandidate;
                            }
                        }

                        if (identifiedRegion == null)
                        {
                            identifiedRegion = address;
                        }

                        string regionClean = identifiedRegion.Trim();
                        regionClean = Regex.Replace(regionClean, @"\s+", " ").Trim();

                        salesDataList.Add(new SatisVerisi
                        {
                            Sehir = regionClean,
                            IlacAdi = drugNameClean,
                            SatilanAdet = soldQuantity
                        });
                    }
                }

                if (salesDataList.Any())
                {
                    var groupedData = salesDataList
                        .GroupBy(s => new { s.IlacAdi, s.Sehir })
                        .Select(g => new
                        {
                            IlacAdi = g.Key.IlacAdi,
                            Sehir = g.Key.Sehir,
                            ToplamAdet = g.Sum(s => s.SatilanAdet)
                        })
                        .ToList();

                    var uniqueDrugs = groupedData.Select(g => g.IlacAdi).Distinct().OrderBy(i => i).ToList();
                    var uniqueCities = groupedData.Select(g => g.Sehir).Distinct().OrderBy(s => s).ToList();

                    using var resultWb = new XLWorkbook();
                    var resultWs = resultWb.Worksheets.Add("Çapraz Satış Cədvəli");

                    resultWs.Cell(1, 1).Value = "Dərman Adı";

                    for (int i = 0; i < uniqueCities.Count; i++)
                    {
                        resultWs.Cell(1, i + 2).Value = uniqueCities[i];
                    }

                    for (int i = 0; i < uniqueDrugs.Count; i++)
                    {
                        string currentDrug = uniqueDrugs[i];
                        int currentRow = i + 2;

                        resultWs.Cell(currentRow, 1).Value = currentDrug;

                        for (int j = 0; j < uniqueCities.Count; j++)
                        {
                            string currentCity = uniqueCities[j];
                            int currentCol = j + 2;

                            var salesRecord = groupedData
                                .FirstOrDefault(g => g.IlacAdi == currentDrug && g.Sehir == currentCity);

                            decimal totalQuantity = salesRecord?.ToplamAdet ?? 0m;

                            resultWs.Cell(currentRow, currentCol).Value = totalQuantity;
                        }
                    }

                    resultWs.Row(1).Style.Font.Bold = true;
                    resultWs.Columns().AdjustToContents();


                    using (var sfd = new SaveFileDialog
                    {
                        Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                        Title = "Nəticəni Excel olaraq Yadda Saxla",
                        FileName = "AvromedTamSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"

                    })

                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            resultWb.SaveAs(sfd.FileName);
                            MessageBox.Show(

                                "Nəticə saxlanıldı:\n" + sfd.FileName,
                                 "Uğurlu",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                             );
                        }
                }
                else
                {
                    MessageBox.Show("Excel faylında emal edilə biləcək etibarlı satış məlumatı tapılmadı.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Bir xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //sedef

            // JSON dosyasının yolu (örneğin uygulama klasöründe Mappings/sedef.json)
            string jsonPath = Path.Combine(Application.StartupPath, "Mappings", "sedef.json");

            Dictionary<string, string> cityMappings;

            if (!File.Exists(jsonPath))
            {
                MessageBox.Show($"JSON dosyası bulunamadı:\n{jsonPath}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string jsonText = File.ReadAllText(jsonPath);
                cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON dosyası okunurken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // ---- Burada totals dictionary kodda kalıyor ----
            Dictionary<string, double> totals = new();

            string[] notCities = new[]
            {
                "ampula", "məhlul", "gel", "tabletka", "şampun", "kapsul", "aerozol", "damcı", "sprey", "krem"
            };
            string[] possibleMedicines = new[]
{
    "Aeromaks", "Aerovin", "Aksomed", "Aqneteks Forte", "Arovaban", "Artron", "Artron A", "Buderen",
    "Dekspan", "Diafleks", "Difluvid", "Efilen", "Egeron", "Elafra", "Enurezin", "Epafor", "Estilak",
    "Flagimet", "Flaksidel", "Foligin-5", "Gera", "Hifes", "Ginestil Lavanda", "Ginestil",
    "Klindabioks", "Lekart", "Mastaq gel", "Mukobronx", "Natamiks", "Neomezol", "Nervio B12",
    "Neyrotilin", "Panorin", "Panorin A", "Papil Derma", "Papil-Off", "Probien", "Proktotrombin",
    "Protesol", "Psilomusil", "Qliaton Forte","Resalfu 25/125", "Resalfu 25/250","Rinoret", "Rumalon", "Rudaza", "Senaval",
    "Serfunal", "Soludazol", "Spazmolizin", "Tromisin", "Ulpriks", "Uroseptin", "Vasklor",
    "Viotiser", "Vitokalsit", "Yenlip",
    "Biostrepta", "Profideks",
    "Skarvis",
    "Bebinorm", "Eribenz", "Flurapid", "Kolonat-TF", "Misopreks", "Vitanur",
    "Meronat TF 1 000", "Meronat TF 500",

};
            string[] possibleCities = new[]
            {
                "Ağcabədi", "Ağdam", "Ağdaş", "Ağstafa", "Ağsu",
                "Astara",
                "Bakı", "Balakən", "Beyləqan", "Bərdə",
                "Biləsuvar",
                "Cəbrayıl", "Cəlilabad",
                "Daşkəsən",
                "Füzuli",
                "Gədəbəy", "Gəncə", "Goranboy", "Göygöl", "Göyçay",
                "Hacıqabul", "Hövsan",
                "İmişli", "İsmayıllı",
                "Kəlbəcər", "Kürdəmir",
                "Laçın", "Lerik", "Lənkəran",
                "Masallı", "Mingəçevir",
                "Naftalan", "Neftçala",
                "Oğuz",
                "Qəbələ", "Qax", "Qazax", "Qobustan", "Quba", "Qubadlı", "Qusar",
                "Saatlı", "Sabirabad", "Sədərək", "Salyan", "Samux", "Şabran", "Şahbuz", "Şamaxı", "Şəki", "Şəmkir", "Şərur", "Şirvan", "Siyəzən", "Sumqayıt",
                "Tərtər", "Tovuz",
                "Ucar",
                "Xaçmaz", "Xızı", "Xocalı", "Xocavənd", "Xudat",
                "Yardımlı", "Yevlax",
                "Zaqatala", "Zəngilan", "Zərdab"
            };

            void AddOrUpdate(string medName, string city, double count)
            {
                if (string.IsNullOrWhiteSpace(city))
                    return;

                string normCity = NormalizeText(city);

                // Şehir eşleme kontrolü JSON'dan geldiği için burada da kullanılır
                if (cityMappings.TryGetValue(normCity, out string mappedCity))
                {
                    city = mappedCity;
                }
                else
                {
                    var match = cityMappings.Keys.FirstOrDefault(k => normCity.Contains(NormalizeText(k)));
                    if (match != null)
                        city = cityMappings[match];
                    else
                        city = char.ToUpper(city[0]) + city.Substring(1).ToLowerInvariant();
                }

                bool isRealCity = possibleCities.Any(c => NormalizeText(c) == NormalizeText(city));
                bool looksLikeNotCity = notCities.Any(f => normCity.Contains(f));

                if (!isRealCity && looksLikeNotCity)
                    return;

                string key = $"{medName.ToLowerInvariant()}|{city.ToLowerInvariant()}";
                if (totals.ContainsKey(key))
                    totals[key] += count;
                else
                    totals[key] = count;
            }

            string NormalizeText(string text)
            {
                return text.ToLowerInvariant()
                    .Replace("ə", "e")
                    .Replace("ı", "i")
                    .Replace("ö", "o")
                    .Replace("ü", "u")
                    .Replace("ç", "c")
                    .Replace("ş", "s")
                    .Replace("ğ", "g");
            }

            var dt = new DataTable();
            dt.Columns.Add("Dərman adı", typeof(string));
            Dictionary<string, int> medicines = new();
            int lastColumn = 0;
            DataTable specialMeds = new();
            specialMeds.Columns.Add("Dərman adı", typeof(string));
            specialMeds.Columns.Add("Şəhər və ya Filial", typeof(string));
            specialMeds.Columns.Add("Say", typeof(double));
            string lastMedName = "";

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Dosyası Seçin"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            foreach (var row in ws.RowsUsed())
            {
                int outline = row.OutlineLevel;
                if (outline == 0)
                {
                    string medName = "";
                    for (int i = 1; i <= 5; i++)
                    {
                        var val = row.Cell(i).GetString().Trim();
                        if (!string.IsNullOrEmpty(val))
                        {
                            foreach (var med in possibleMedicines)
                            {
                                if (val.ToLower().Contains(med.ToLower()))
                                {
                                    medName = val;
                                    break;
                                }
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(medName))
                            break;
                    }
                    if (!string.IsNullOrWhiteSpace(medName))
                        lastMedName = medName;
                    else
                        lastMedName = "";
                }

                if (outline == 1 || !string.IsNullOrWhiteSpace(lastMedName))
                {
                    string rawPharmacyInfo = row.Cell(2).GetString().Trim();
                    double medCount = row.Cell(3).GetValue<double?>() ?? 0;

                    if (medCount == 0 || string.IsNullOrWhiteSpace(rawPharmacyInfo))
                        continue;

                    string extractedCity = "";
                    int bestPos = int.MaxValue;
                    string normPharmacy = NormalizeText(rawPharmacyInfo);

                    foreach (var city in possibleCities)
                    {
                        string normCity = NormalizeText(city);
                        int pos = normPharmacy.IndexOf(normCity);
                        if (pos >= 0 && pos < bestPos)
                        {
                            bestPos = pos;
                            extractedCity = city;
                        }
                    }

                    string location;
                    if (!string.IsNullOrEmpty(extractedCity))
                    {
                        location = extractedCity;
                    }
                    else
                    {
                        bool isMedicine = possibleMedicines.Any(med =>
                            NormalizeText(rawPharmacyInfo).Contains(NormalizeText(med)));

                        if (isMedicine)
                            continue;

                        location = rawPharmacyInfo.Trim();
                    }

                    AddOrUpdate(lastMedName, location, medCount);
                }
            }

            // --- Sonraki tablo ve dosya kaydetme kısmı aynı kalıyor ---

            DataTable finalTable = new();
            finalTable.Columns.Add("Dərman adı", typeof(string));
            finalTable.Columns.Add("Şəhər və ya Filial", typeof(string));
            finalTable.Columns.Add("Say", typeof(double));

            foreach (var kvp in totals)
            {
                var parts = kvp.Key.Split('|');
                var row = finalTable.NewRow();
                row["Dərman adı"] = parts[0];
                row["Şəhər və ya Filial"] = parts[1];
                row["Say"] = kvp.Value;
                finalTable.Rows.Add(row);
            }

            DataTable pivotTable = new();
            pivotTable.Columns.Add("Dərman adı", typeof(string));

            var uniqueLocations = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Şəhər və ya Filial"))
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            foreach (var location in uniqueLocations)
                pivotTable.Columns.Add(location, typeof(double));

            var uniqueMeds = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Dərman adı"))
                .Distinct()
                .OrderBy(m => m)
                .ToList();

            foreach (var med in uniqueMeds)
            {
                var newRow = pivotTable.NewRow();
                newRow["Dərman adı"] = med;

                foreach (var location in uniqueLocations)
                {
                    var match = finalTable.AsEnumerable()
                        .FirstOrDefault(r => r.Field<string>("Dərman adı") == med && r.Field<string>("Şəhər və ya Filial") == location);

                    newRow[location] = match != null ? match.Field<double>("Say") : 0;
                }

                pivotTable.Rows.Add(newRow);
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "SedefSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using var saveWb = new XLWorkbook();
                    var ws1 = saveWb.Worksheets.Add("Pivot_Hesabatı");

                    for (int c = 0; c < pivotTable.Columns.Count; c++)
                    {
                        ws1.Cell(1, c + 1).Value = pivotTable.Columns[c].ColumnName;
                    }

                    for (int r = 0; r < pivotTable.Rows.Count; r++)
                    {
                        for (int c = 0; c < pivotTable.Columns.Count; c++)
                        {
                            var cell = ws1.Cell(r + 2, c + 1);
                            var value = pivotTable.Rows[r][c];

                            if (value is double d)
                            {
                                cell.Value = d;
                            }
                            else
                            {
                                cell.Value = value?.ToString() ?? "";
                            }
                        }
                    }

                    ws1.Columns().AdjustToContents();
                    saveWb.SaveAs(sfd.FileName);

                    MessageBox.Show(
                        "Nəticə saxlanıldı:\n" + sfd.FileName,
                        "Uğurlu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //azerimed


            // JSON dosyasının yolu (örnek: uygulama klasöründe Mappings/azerimed.json)
            string jsonPath = Path.Combine(Application.StartupPath, "Mappings", "azerimed.json");

            if (!File.Exists(jsonPath))
            {
                MessageBox.Show($"JSON dosyası bulunamadı:\n{jsonPath}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Dictionary<string, string> cityMappings;
            try
            {
                string jsonText = File.ReadAllText(jsonPath);
                cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON dosyası okunurken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Faylını Seç"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            string Normalize(string input)
            {
                return input.Trim().ToLower()
                    .Replace("ə", "e")
                    .Replace("ı", "i")
                    .Replace("ö", "o")
                    .Replace("ü", "u")
                    .Replace("ç", "c")
                    .Replace("ş", "s")
                    .Replace("ğ", "g");
            }

            var pivot = new Dictionary<string, Dictionary<string, double>>();

            int firstRow = ws.FirstRowUsed().RowNumber() + 1;
            int lastRow = ws.LastRowUsed().RowNumber();

            for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
            {
                var row = ws.Row(rowNum);

                string eraziRaw = row.Cell(2).GetString();
                string seher = Normalize(eraziRaw.Split('|')[0]);

                // JSON'dan şehir eşleme (tam eşleşme)
                if (cityMappings.TryGetValue(seher, out string mappedCity))
                {
                    seher = mappedCity;
                }
                else
                {
                    // Eğer tam eşleşme yoksa, içerik kontrolü ile eşleme yap
                    foreach (var kv in cityMappings)
                    {
                        if (seher.Contains(kv.Key))
                        {
                            seher = kv.Value;
                            break;
                        }
                    }
                }

                string malAdi = Normalize(row.Cell(4).GetString());

                if (malAdi.Contains("AEROMAX 50ml"))
                {
                    malAdi = "Aeromax";
                }


                double miqdar = 0;
                double.TryParse(row.Cell(7).GetValue<string>(), out miqdar);

                if (!pivot.ContainsKey(malAdi))
                    pivot[malAdi] = new Dictionary<string, double>();

                if (!pivot[malAdi].ContainsKey(seher))
                    pivot[malAdi][seher] = 0;

                if (miqdar < 0)
                    pivot[malAdi][seher] -= Math.Abs(miqdar);
                else
                    pivot[malAdi][seher] += miqdar;
            }

            using var newWb = new XLWorkbook();
            var newWs = newWb.Worksheets.Add("Pivot");

            var allCities = pivot.SelectMany(p => p.Value.Keys).Distinct().OrderBy(x => x).ToList();

            for (int i = 0; i < allCities.Count; i++)
                newWs.Cell(1, i + 2).Value = allCities[i];

            int rowIndex = 2;
            foreach (var mal in pivot.Keys.OrderBy(x => x))
            {
                newWs.Cell(rowIndex, 1).Value = mal;

                for (int colIndex = 0; colIndex < allCities.Count; colIndex++)
                {
                    string city = allCities[colIndex];
                    double value = pivot[mal].ContainsKey(city) ? pivot[mal][city] : 0;
                    newWs.Cell(rowIndex, colIndex + 2).Value = value;
                }

                rowIndex++;
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "AzerimedSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    newWb.SaveAs(sfd.FileName);

                MessageBox.Show(
                    "Nəticə saxlanıldı:\n" + sfd.FileName,
                    "Uğurlu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //zeytun aptek

            // Normalize Med Name
            string NormalizeMedName(string raw)
            {
                int lastStart = raw.LastIndexOf("(");
                int lastEnd = raw.LastIndexOf(")");
                if (lastStart >= 0 && lastEnd > lastStart)
                    raw = raw.Remove(lastStart, lastEnd - lastStart + 1);

                if (raw.Contains("№"))
                {
                    raw = raw.Substring(0, raw.IndexOf("№")).Trim();
                }

                return raw.Trim().ToLowerInvariant();
            }

            // Normalize text for city name matching
            string NormalizeText(string text)
            {
                return text.ToLowerInvariant()
                    .Replace("ə", "e")
                    .Replace("ı", "i")
                    .Replace("ö", "o")
                    .Replace("ü", "u")
                    .Replace("ç", "c")
                    .Replace("ş", "s")
                    .Replace("ğ", "g");
            }

            // JSON mapping dosyasını oku
            string mappingJsonPath = Path.Combine(Application.StartupPath, "mappings", "zeytun.json");
            Dictionary<string, string> cityMappings = new Dictionary<string, string>();

            if (File.Exists(mappingJsonPath))
            {
                var jsonText = File.ReadAllText(mappingJsonPath);
                cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);
                cityMappings = cityMappings.ToDictionary(kvp => NormalizeText(kvp.Key), kvp => kvp.Value);
            }
            else
            {
                MessageBox.Show("Mapping dosyası bulunamadı: " + mappingJsonPath, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Zeytun excel faylını seçin"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            var dt = new DataTable();
            dt.Columns.Add("Dərman adı", typeof(string));
            Dictionary<string, int> medicines = new Dictionary<string, int>();
            int lastColumn = 0;
            bool isSpecial = false;
            DataTable specialMeds = new DataTable();
            specialMeds.Columns.Add("Dərman adı", typeof(string));
            specialMeds.Columns.Add("Şəhər və ya Filial", typeof(string));
            specialMeds.Columns.Add("Say", typeof(double));
            string lastBranchName = "";
            string lastMedName = "";

            foreach (var row in ws.RowsUsed())
            {
                int outline = row.OutlineLevel;

                // Filial adını yakala
                for (int i = 1; i <= 5; i++)
                {
                    var val = row.Cell(i).GetString().Trim();
                    if (!string.IsNullOrEmpty(val) && val.ToLower().Contains("filial"))
                    {
                        lastBranchName = val;
                        break;
                    }
                }

                if (string.IsNullOrWhiteSpace(lastBranchName))
                    lastBranchName = "Bilinməyən Filial";

                if (outline == 1)
                {
                    string cityNameRaw = row.Cell(2).GetString().Trim();
                    string normalizedCity = NormalizeText(cityNameRaw);
                    string cityName;

                    if (cityMappings.TryGetValue(normalizedCity, out var mappedCity))
                    {
                        cityName = mappedCity;
                    }
                    else
                    {
                        cityName = cityNameRaw;
                    }

                    if (cityName.Contains("Şəxsi apteklər"))
                    {
                        isSpecial = true;
                    }
                    else
                    {
                        if (!dt.Columns.Contains(cityName))
                        {
                            dt.Columns.Add(cityName, typeof(double));
                        }
                        lastColumn = dt.Columns.IndexOf(cityName);
                        isSpecial = false;
                    }
                }
                else if (outline == 2)
                {
                    string medNameRaw = row.Cell(2).GetString().Trim();

                    if (isSpecial)
                    {
                        lastMedName = medNameRaw;

                        double medCountSpecial = row.Cell(3).GetValue<double?>() ?? 0;
                        if (medCountSpecial != 0)
                        {
                            var newRow = specialMeds.NewRow();
                            newRow["Dərman adı"] = NormalizeMedName(medNameRaw);
                            newRow["Şəhər və ya Filial"] = lastBranchName;
                            newRow["Say"] = medCountSpecial;
                            specialMeds.Rows.Add(newRow);
                        }
                        continue;
                    }

                    string medName = NormalizeMedName(medNameRaw);
                    double medCount = row.Cell(3).GetValue<double?>() ?? 0;

                    if (!medicines.ContainsKey(medName))
                    {
                        var dataRow = dt.NewRow();
                        dataRow[0] = medName;
                        for (int i = 1; i < dt.Columns.Count; i++)
                            dataRow[i] = 0;

                        dataRow[lastColumn] = medCount;
                        dt.Rows.Add(dataRow);
                        medicines.Add(medName, dt.Rows.Count - 1);
                    }
                    else
                    {
                        var dr = dt.Rows[medicines[medName]];
                        double currentVal = double.TryParse(dr[lastColumn]?.ToString(), out var cur) ? cur : 0;
                        dr[lastColumn] = currentVal + medCount;
                    }
                }
            }

            DataTable finalTable = new DataTable();
            finalTable.Columns.Add("Dərman adı", typeof(string));
            finalTable.Columns.Add("Şəhər və ya Filial", typeof(string));
            finalTable.Columns.Add("Say", typeof(double));

            string[] notCities = new[] { "ampula", "məhlul", "gel", "tabletka", "şampun", "kapsul", "aerozol", "damcı", "sprey", "krem" };
            Dictionary<string, double> totals = new Dictionary<string, double>();

            void AddOrUpdate(string medName, string city, double count)
            {
                if (!string.IsNullOrWhiteSpace(city))
                    city = char.ToUpper(city[0]) + city[1..].ToLowerInvariant();

                if (string.IsNullOrWhiteSpace(city) || notCities.Any(f => city.ToLowerInvariant().Contains(f)))
                    return;

                string key = $"{medName.ToLowerInvariant()}|{city.ToLowerInvariant()}";
                if (totals.ContainsKey(key))
                    totals[key] += count;
                else
                    totals[key] = count;
            }

            foreach (DataRow dr in dt.Rows)
            {
                string medName = dr[0].ToString();
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    string city = dt.Columns[i].ColumnName;
                    double count = double.TryParse(dr[i]?.ToString(), out var c) ? c : 0;
                    if (count != 0)
                        AddOrUpdate(medName, city, count);
                }
            }

            foreach (DataRow dr in specialMeds.Rows)
            {
                string medName = dr[0].ToString();
                string city = dr[1].ToString();
                double count = double.TryParse(dr[2]?.ToString(), out var c) ? c : 0;
                if (count != 0)
                    AddOrUpdate(medName, city, count);
            }

            foreach (var kvp in totals)
            {
                var parts = kvp.Key.Split('|');
                var row = finalTable.NewRow();
                row["Dərman adı"] = parts[0];
                row["Şəhər və ya Filial"] = parts[1];
                row["Say"] = kvp.Value;
                finalTable.Rows.Add(row);
            }

            DataTable pivotTable = new DataTable();
            pivotTable.Columns.Add("Dərman adı", typeof(string));

            var uniqueLocations = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Şəhər və ya Filial"))
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            foreach (var location in uniqueLocations)
                pivotTable.Columns.Add(location, typeof(double));

            var uniqueMeds = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Dərman adı"))
                .Distinct()
                .OrderBy(m => m)
                .ToList();

            foreach (var med in uniqueMeds)
            {
                var newRow = pivotTable.NewRow();
                newRow["Dərman adı"] = med;

                foreach (var location in uniqueLocations)
                {
                    var match = finalTable.AsEnumerable()
                        .FirstOrDefault(r => r.Field<string>("Dərman adı") == med && r.Field<string>("Şəhər və ya Filial") == location);

                    newRow[location] = match != null ? match.Field<double>("Say") : 0;
                }

                pivotTable.Rows.Add(newRow);
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "ZeytunSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using var saveWb = new XLWorkbook();
                    var ws1 = saveWb.Worksheets.Add("Pivot_Hesabatı");

                    for (int c = 0; c < pivotTable.Columns.Count; c++)
                    {
                        ws1.Cell(1, c + 1).Value = pivotTable.Columns[c].ColumnName;
                    }

                    for (int r = 0; r < pivotTable.Rows.Count; r++)
                    {
                        for (int c = 0; c < pivotTable.Columns.Count; c++)
                        {
                            var cell = ws1.Cell(r + 2, c + 1);
                            var value = pivotTable.Rows[r][c];

                            if (value is double d)
                            {
                                cell.Value = d;
                            }
                            else
                            {
                                cell.Value = value?.ToString() ?? "";
                            }
                        }
                    }

                    ws1.Columns().AdjustToContents();
                    saveWb.SaveAs(sfd.FileName);

                    MessageBox.Show(
                        "Nəticə saxlanıldı:\n" + sfd.FileName,
                        "Uğurlu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }

            dataGridView1.DataSource = pivotTable;
            dataGridView1.Visible = false;
            dataGridView1.SendToBack();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            //radez rehim
            // JSON mapping faylının yolu
            string jsonPath = Path.Combine(Application.StartupPath, "Mappings", "radezrehim.json");

            if (!File.Exists(jsonPath))
            {
                MessageBox.Show($"JSON dosyası tapılmadı:\n{jsonPath}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Dictionary<string, string> cityMappings;
            try
            {
                string jsonText = File.ReadAllText(jsonPath);
                cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON oxunarkən xəta:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Faylını Seç"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            // Yeni normalize fonksiyonu: karakterleri ve birden fazla boşluğu temizler
            string Normalize(string input)
            {
                if (string.IsNullOrEmpty(input))
                    return string.Empty;

                // Birden fazla boşluğu tek boşluğa indirmek için Regex kullan
                string trimmed = Regex.Replace(input.Trim(), @"\s+", " ");

                return trimmed.ToLower()
                    .Replace("ə", "e")
                    .Replace("ı", "i")
                    .Replace("ö", "o")
                    .Replace("ü", "u")
                    .Replace("ç", "c")
                    .Replace("ş", "s")
                    .Replace("ğ", "g");
            }

            // JSON anahtarlarını önceden normalize edip yeni bir sözlüğe kaydediyoruz
            var normalizedCityMappings = new Dictionary<string, string>();
            foreach (var kv in cityMappings)
            {
                normalizedCityMappings[Normalize(kv.Key)] = kv.Value;
            }

            var pivot = new Dictionary<string, Dictionary<string, double>>();

            int firstRow = ws.FirstRowUsed().RowNumber() + 1; // başlıqdan sonrakı sətir
            int lastRow = ws.LastRowUsed().RowNumber();

            for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
            {
                var row = ws.Row(rowNum);

                string musteriRaw = row.Cell(5).GetString(); // Müştəri (aptek adı)
                string normalizedMusteri = Normalize(musteriRaw);

                string seher = musteriRaw; // default olaraq orijinal

                // Normalize edilmiş JSON anahtarlarıyla karşılaştırma yap
                if (normalizedCityMappings.ContainsKey(normalizedMusteri))
                {
                    seher = normalizedCityMappings[normalizedMusteri];
                }
                else
                {
                    // Eğer tam eşleşme yoksa, contains metoduyla arama yap
                    foreach (var kv in normalizedCityMappings)
                    {
                        if (normalizedMusteri.Contains(kv.Key))
                        {
                            seher = kv.Value;
                            break;
                        }
                    }
                }

                string malAdi = Normalize(row.Cell(3).GetString()); // Preparat
                double miqdar = 0;
                double.TryParse(row.Cell(6).GetValue<string>(), out miqdar);

                if (!pivot.ContainsKey(malAdi))
                    pivot[malAdi] = new Dictionary<string, double>();

                if (!pivot[malAdi].ContainsKey(seher))
                    pivot[malAdi][seher] = 0;

                pivot[malAdi][seher] += miqdar;
            }

            // Yeni Excel faylı yaradılır
            using var newWb = new XLWorkbook();
            var newWs = newWb.Worksheets.Add("Pivot");

            var allCities = pivot.SelectMany(p => p.Value.Keys).Distinct().OrderBy(x => x).ToList();

            // başlıqlar (şəhərlər)
            for (int i = 0; i < allCities.Count; i++)
                newWs.Cell(1, i + 2).Value = allCities[i];

            int rowIndex = 2;
            foreach (var mal in pivot.Keys.OrderBy(x => x))
            {
                newWs.Cell(rowIndex, 1).Value = mal;

                for (int colIndex = 0; colIndex < allCities.Count; colIndex++)
                {
                    string city = allCities[colIndex];
                    double value = pivot[mal].ContainsKey(city) ? pivot[mal][city] : 0;
                    newWs.Cell(rowIndex, colIndex + 2).Value = value;
                }

                rowIndex++;
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "RadezRehimSatishHesabati" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    newWb.SaveAs(sfd.FileName);

                MessageBox.Show(
                        "Nəticə saxlanıldı:\n" + sfd.FileName,
                        "Uğurlu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {

            //sonar

            // JSON mapping faylının yolu
            string jsonPath = Path.Combine(Application.StartupPath, "Mappings", "sonar.json");

            if (!File.Exists(jsonPath))
            {
                MessageBox.Show($"JSON faylı tapılmadı:\n{jsonPath}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Dictionary<string, string> cityMappings;
            try
            {
                string jsonText = File.ReadAllText(jsonPath);
                cityMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON oxunarkən xəta:\n" + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Faylını Seç"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            string Normalize(string input)
            {
                if (string.IsNullOrEmpty(input))
                    return string.Empty;
                string trimmed = Regex.Replace(input.Trim(), @"\s+", " ");
                return trimmed.ToLower()
                    .Replace("ə", "e").Replace("ı", "i").Replace("ö", "o")
                    .Replace("ü", "u").Replace("ç", "c").Replace("ş", "s").Replace("ğ", "g");
            }

            var normalizedCityMappings = new Dictionary<string, string>();
            foreach (var kv in cityMappings)
            {
                normalizedCityMappings[Normalize(kv.Key)] = kv.Value;
            }

            var pivot = new Dictionary<string, Dictionary<string, double>>();

            var firstRow = ws.FirstRowUsed();
            var libartCol = firstRow.CellsUsed().FirstOrDefault(c => Normalize(c.Value.ToString()) == "libart")?.Address.ColumnNumber;
            var cityCol = firstRow.CellsUsed().FirstOrDefault(c => Normalize(c.Value.ToString()) == "city")?.Address.ColumnNumber;
            var quanCol = firstRow.CellsUsed().FirstOrDefault(c => Normalize(c.Value.ToString()) == "quan")?.Address.ColumnNumber;

            // Yalnız city, libart ve quan sütunlarını istifadə edirik
            if (libartCol == null || cityCol == null || quanCol == null)
            {
                MessageBox.Show("Lazımi sütunlar (libart, city, quan) tapılmadı.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int firstDataRow = firstRow.RowNumber() + 1;
            int lastDataRow = ws.LastRowUsed().RowNumber();

            for (int rowNum = firstDataRow; rowNum <= lastDataRow; rowNum++)
            {
                var row = ws.Row(rowNum);

                string libart = row.Cell(libartCol.Value).GetString();
                string cityRaw = row.Cell(cityCol.Value).GetString();
                double quan = 0;
                double.TryParse(row.Cell(quanCol.Value).GetValue<string>(), out quan);

                string normalizedCity = Normalize(cityRaw);
                string finalCity = cityRaw;

                // "city" sütunundakı dəyəri JSON mappinginə uyğunlaşdırırıq
                if (normalizedCityMappings.ContainsKey(normalizedCity))
                {
                    finalCity = normalizedCityMappings[normalizedCity];
                }

                // Digər şəhərlər (masazir, berde və s.) üçün ayrı mappinglər əlavə edə bilərsiniz
                if (!pivot.ContainsKey(libart))
                    pivot[libart] = new Dictionary<string, double>();

                if (!pivot[libart].ContainsKey(finalCity))
                    pivot[libart][finalCity] = 0;

                pivot[libart][finalCity] += quan;
            }

            // Yeni Excel faylı yaradılır
            using var newWb = new XLWorkbook();
            var newWs = newWb.Worksheets.Add("Pivot");

            var allCities = pivot.SelectMany(p => p.Value.Keys).Distinct().OrderBy(x => x).ToList();

            newWs.Cell(1, 1).Value = "libart";
            for (int i = 0; i < allCities.Count; i++)
                newWs.Cell(1, i + 2).Value = allCities[i];

            int rowIndex = 2;
            foreach (var mal in pivot.Keys.OrderBy(x => x))
            {
                newWs.Cell(rowIndex, 1).Value = mal;

                for (int colIndex = 0; colIndex < allCities.Count; colIndex++)
                {
                    string city = allCities[colIndex];
                    double value = pivot[mal].ContainsKey(city) ? pivot[mal][city] : 0;
                    newWs.Cell(rowIndex, colIndex + 2).Value = value;
                }

                rowIndex++;
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "SonarElnurSatishHesabati" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    newWb.SaveAs(sfd.FileName);

                MessageBox.Show(
                        "Nəticə saxlanıldı:\n" + sfd.FileName,
                        "Uğurlu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
            }
        }
    }
}