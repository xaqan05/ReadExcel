using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

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

            using var ofd = new OpenFileDialog()
            {
                Filter = "Excel Dosyaları|*.xlsx;*.xls",
                Title = "Excel Dosyası Seçin"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            using var wb = new XLWorkbook(ofd.FileName);
            var ws = wb.Worksheet(1);

            var dt = new DataTable();
            dt.Columns.Add("Dərman adı", typeof(string));
            List<string> prevCities = [];
            Dictionary<string, int> medicines = [];
            int lastColumn = 0;
            bool isSpecial = false;
            DataTable specialMeds = new();
            specialMeds.Columns.Add("Dərman adı", typeof(string));
            specialMeds.Columns.Add("Şəhər və ya Filial", typeof(string));
            specialMeds.Columns.Add("Say", typeof(double));
            string lastBranchName = "";
            string lastMedName = "";

            foreach (var row in ws.RowsUsed())
            {
                int outline = row.OutlineLevel;

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

                    string cityName = row.Cell(2).GetString();

                    if (cityName.Contains("Şəxsi apteklər"))
                    {
                        isSpecial = true;
                    }
                    else
                    {
                        dt.Columns.Add(cityName, typeof(double));
                        lastColumn = dt.Columns.Count - 1;
                        isSpecial = false;
                    }
                }
                else if (outline == 2)
                {
                    string medNameRaw = row.Cell(2).GetString().Trim();
                    if (isSpecial)
                    {
                        lastMedName = medNameRaw;
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
                else if (outline == 3 && isSpecial)
                {

                    string rawMedName = lastMedName;
                    string rawPharmacyInfo = row.Cell(2).GetString().Trim();

                    double count = row.Cell(3).GetValue<double?>() ?? 0;

                    string medName = NormalizeMedName(rawMedName);

                    string[] possibleCities = new[] {
                        "Ağcabədi", "Ağdam", "Beyləqan", "Bərdə", "Füzuli", "Tərtər",
                        "Gədəbəy", "Gəncə", "Goranboy", "Naftalan", "Şəmkir-Çinarlı", "Şəmkir",
                        "Astara", "Biləsuvar", "Cəlilabad", "Lerik", "Lənkəran", "Masallı", "Yardımlı",
                        "Ağdaş", "Göyçay", "Mingəçevir", "Ucar", "Yevlax", "Zərdab",
                        "Alıcılar",
                        "Ağstafa", "Qazax", "Tovuz",
                        "Balakən", "İsmayıllı", "Qax", "Qəbələ", "Şəki", "Zaqatala","Oğuz Xaçmaz kəndi",
                        "Hacıqabul", "İmişli", "Kürdəmir", "Neftçala", "Saatlı", "Sabirabad", "Salyan", "Şamaxı", "Şirvan",
                        "Quba", "Qusar", "Şabran", "Siyezen", "Xaçmaz", "Xudat"
                    };

                    string normPharmacy = NormalizeText(rawPharmacyInfo);
                    string extractedCity = "";
                    int bestPos = int.MaxValue;

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
                        location = $"{lastBranchName.Trim()} - {rawPharmacyInfo.Trim()}";
                    }

                    var dr = specialMeds.NewRow();
                    dr[0] = medName;
                    dr[1] = location;
                    dr[2] = count;
                    specialMeds.Rows.Add(dr);
                }
            }

            DataTable finalTable = new();
            finalTable.Columns.Add("Dərman adı", typeof(string));
            finalTable.Columns.Add("Şəhər və ya Filial", typeof(string));
            finalTable.Columns.Add("Say", typeof(double));

            string[] notCities = ["ampula", "məhlul", "gel", "tabletka", "şampun", "kapsul", "aerozol", "damcı", "sprey", "krem"];
            Dictionary<string, double> totals = new();

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
                FileName = "AvromedRegionSatish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
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

        private void button2_Click(object sender, EventArgs e)
        {
            string[] notCities = ["ampula", "məhlul", "gel", "tabletka", "şampun", "kapsul", "aerozol", "damcı", "sprey", "krem"];
            string[] possibleMedicines = new[]
            {
            "Aerovin", "Aksomed", "Aqneteks Forte", "Artron", "Artron A", "Biostrepta", "Buderen", "Dekspan",
            "Dermizol-G", "Diafleks", "Efilen", "Egeron", "Elafra", "Epafor", "Eribenz", "Estilak",
            "Flagimet", "Flaksidel", "Foligin-5", "Gera", "Hifes", "Lekart", "Mastaq gel", "Misopreks",
            "Mukobronx", "Natamiks", "Neomezol", "Nervio B12", "Neyrotilin", "Panorin", "Panorin A", "Papil Derma",
            "Papil-Off", "Probien", "Protesol", "Qliaton Forte", "Resalfu", "Rinoret", "Rudaza", "Senaval",
            "Skarvis", "Spazmolizin", "Tromisin", "Ulpriks", "Uroseptin", "Vasklor", "Viotiser", "Vitokalsit", "Yenlip"
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

            Dictionary<string, double> totals = new();

            void AddOrUpdate(string medName, string city, double count)
            {
                if (!string.IsNullOrWhiteSpace(city))
                    city = char.ToUpper(city[0]) + city[1..].ToLowerInvariant();

                bool isRealCity = possibleCities.Contains(city);
                bool looksLikeNotCity = notCities.Any(f => city.ToLowerInvariant().Contains(f));

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
            Dictionary<string, int> medicines = [];
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
    }
}