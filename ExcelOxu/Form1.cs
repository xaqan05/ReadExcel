using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;

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
                while (raw.Contains("(") && raw.Contains(")"))
                {
                    int start = raw.IndexOf("(");
                    int end = raw.IndexOf(")", start);
                    if (start >= 0 && end > start)
                        raw = raw.Remove(start, end - start + 1);
                    else
                        break;
                }

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
            specialMeds.Columns.Add("Şəhər", typeof(string));
            specialMeds.Columns.Add("Say", typeof(double));
            string lastMedName = "";

            foreach (var row in ws.RowsUsed())
            {
                int outline = row.OutlineLevel;

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

                    string extractedCity = "";
                    string[] possibleCities = new[]
{
     "Ağcabədi", "Ağdam", "Ağdaş", "Ağstafa", "Ağsu", "Astara", "Balakən", "Bərdə",
    "Beyləqan", "Biləsuvar", "Cəbrayıl", "Cəlilabad", "Daşkəsən", "Füzuli", "Gədəbəy", "Gəncə",
    "Goranboy", "Göyçay", "Göygöl", "Hacıqabul", "İmişli", "İsmayıllı", "Kəlbəcər", "Kürdəmir",
    "Qax", "Qazax", "Qəbələ", "Qobustan", "Quba", "Qubadlı", "Qusar", "Laçın", "Lənkəran",
    "Lerik", "Masallı", "Mingəçevir", "Naftalan", "Naxçıvan", "Neftçala", "Oğuz", "Ordubad",
    "Saatlı", "Sabirabad", "Şirvan", "Şabran", "Şahbuz", "Şəki", "Salyan", "Şamaxı", "Şəmkir",
    "Samux", "Şərur", "Siyəzən", "Sumqayıt", "Şuşa", "Tərtər", "Tovuz", "Ucar", "Xaçmaz", "Xankəndi",
    "Xızı", "Xocalı", "Xocavənd", "Yardımlı", "Yevlax", "Zaqatala", "Zəngilan", "Zərdab"
};



                    string normPharmacy = NormalizeText(rawPharmacyInfo);

                    foreach (var city in possibleCities)
                    {
                        string normCity = NormalizeText(city);
                        if (normPharmacy.Contains(normCity))
                        {
                            extractedCity = city;
                            break;
                        }
                        else
                        {
                            Debug.WriteLine($"Tapılmadı: {rawPharmacyInfo}");
                        }

                    }


                    if (!string.IsNullOrEmpty(extractedCity))
                    {
                        var dr = specialMeds.NewRow();
                        dr[0] = medName;
                        dr[1] = extractedCity;
                        dr[2] = count;
                        specialMeds.Rows.Add(dr);
                    }
                    else
                    {
                        var dr = specialMeds.NewRow();
                        dr[0] = medName;
                        dr[1] = rawPharmacyInfo; 
                        dr[2] = count;
                        specialMeds.Rows.Add(dr);
                    }


                }
            }

            DataTable finalTable = new();
            finalTable.Columns.Add("Dərman adı", typeof(string));
            finalTable.Columns.Add("Şəhər", typeof(string));
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
                row["Şəhər"] = parts[1];
                row["Say"] = kvp.Value;
                finalTable.Rows.Add(row);
            }

            DataTable pivotTable = new();
            pivotTable.Columns.Add("Dərman adı", typeof(string));

            var uniqueCities = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Şəhər"))
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            foreach (var city in uniqueCities)
                pivotTable.Columns.Add(city, typeof(double));

            var uniqueMeds = finalTable.AsEnumerable()
                .Select(r => r.Field<string>("Dərman adı"))
                .Distinct()
                .OrderBy(m => m)
                .ToList();

            foreach (var med in uniqueMeds)
            {
                var newRow = pivotTable.NewRow();
                newRow["Dərman adı"] = med;

                foreach (var city in uniqueCities)
                {
                    var match = finalTable.AsEnumerable()
                        .FirstOrDefault(r => r.Field<string>("Dərman adı") == med && r.Field<string>("Şəhər") == city);

                    newRow[city] = match != null ? match.Field<double>("Say") : 0;
                }

                pivotTable.Rows.Add(newRow);
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Dosyası (*.xlsx)|*.xlsx",
                Title = "Nəticəni Excel olaraq Yadda Saxla",
                FileName = "Satish_Hesabati_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
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
        }
    }
}
