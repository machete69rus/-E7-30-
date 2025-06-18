using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization;
using System.Windows.Forms.DataVisualization.Charting;


namespace WindowsFormsApp2
{

    public partial class Form1 : Form
    {

        private class RawFileData
        {
            public string FileName { get; set; }
            public List<string> Lines { get; set; } = new List<string>();
            public DeviceType? DetectedDevice { get; set; } // Привязка к конкретному прибору
            public string[] Headers { get; set; }
        }

        private Dictionary<string, RawFileData> rawFiles = new Dictionary<string, RawFileData>();
        private Dictionary<string, List<MeasurementRow>> allMeasurements = new Dictionary<string, List<MeasurementRow>>();
        private Dictionary<string, List<(double f, double eps)>> allEpsilons = new Dictionary<string, List<(double, double)>>();

        private TextBox renameTextBox;
        private TabPage renamingTabPage = null;

        private readonly Dictionary<DeviceType, string[]> deviceColumnHeaders = new Dictionary<DeviceType, string[]>
{
    { DeviceType.Вектор, new[] { "Частота f", "Энергия Q", "tg", "Импеданс Im", "Фаза Phase", "Индуктивность L", "Ёмкость C", "Резистивность R" } },
    { DeviceType.Е7_30, new[] { "F[Hz]", "C[F]", "L[H]", "R[Ohm]", "G[S]", "B[S]", "X[Ohm]", "Z[Ohm]", "D[]", "Q[]", "Phi[degree]", "Ub[V]" } }
};

        private readonly string[] TgHeaders = { "D[]", "D []", "tg", "tg(delta)", "tanδ", "tan(delta)" };

        private static void RenameDictionaryKey<TKey, TValue>(
        IDictionary<TKey, TValue> dict, TKey oldKey, TKey newKey)
        {
            if (dict.ContainsKey(oldKey) && !dict.ContainsKey(newKey))
            {
                var val = dict[oldKey];
                dict.Remove(oldKey);
                dict[newKey] = val;
            }
        }

        private enum DeviceType
        {
            Вектор,
            Е7_30
        }

        private DeviceType CurrentDevice
        {
            get
            {
                string selected = cmbDevice.SelectedItem?.ToString();
                if (selected == "Вектор")
                    return DeviceType.Вектор;
                else if (selected == "Е7-30")
                    return DeviceType.Е7_30;
                else
                    return DeviceType.Вектор; // по умолчанию
            }
        }

        private void ParseAndDisplay(string fileName)
        {
            if (!rawFiles.ContainsKey(fileName))
                return;

            var fileData = rawFiles[fileName];
            var lines = fileData.Lines;
            var headers = fileData.Headers;

            string tabTitle = MakeTabTitle(fileName);

            TabPage tabPage = new TabPage(tabTitle)
            {
                Tag = fileName // <--- это обязательно, иначе oldKey будет null
}
            ;
            DataGridView dataGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                ScrollBars = ScrollBars.Both,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            };

            dataGrid.Columns.Clear();
            foreach (var header in headers)
            {
                dataGrid.Columns.Add(header, header);
            }

            foreach (var line in lines)
            {
                var parts = line.Split('\t');
                if (parts.Length != headers.Length) continue;
                dataGrid.Rows.Add(parts);
            }

            tabPage.Controls.Add(dataGrid);
            tabControlFiles.TabPages.Add(tabPage);
        }
        
        public Form1()
        {
            InitializeComponent();

        }

        private void LoadFileToTab(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string tabTitle = MakeTabTitle(fileName);
            var lines = File.ReadAllLines(filePath).ToList();

            if (lines.Count == 0)
                return;

            // Определим заголовки
            var headerLine = lines[0];
            var headers = headerLine.Split('\t');

            // Удалим первую строку для чистых данных
            lines.RemoveAt(0);

            // Определим прибор
            DeviceType? detectedDevice = null;
            foreach (var pair in deviceColumnHeaders)
            {
                if (headers.Length == pair.Value.Length && headers.SequenceEqual(pair.Value))
                {
                    detectedDevice = pair.Key;
                    break;
                }
            }

            if (detectedDevice == null)
            {
                MessageBox.Show($"Файл '{fileName}' не соответствует ни одному поддерживаемому прибору", "Неверный формат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Сохраняем данные
            rawFiles[fileName] = new RawFileData
            {
                FileName = fileName,
                Lines = lines,
                Headers = headers,
                DetectedDevice = detectedDevice
            };

            // Если прибор выбран верно — отображаем
            if (CurrentDevice == detectedDevice)
            {
                ParseAndDisplay(fileName);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbDevice.Items.Add("Вектор");
            cmbDevice.Items.Add("Е7-30");
            cmbDevice.SelectedIndex = 1; // По умолчанию "Вектор"
        }
        
        private void btnLoadFile_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Text files (*.txt)|*.txt";
                openFileDialog.Title = "Выберите один или несколько файлов";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (var file in openFileDialog.FileNames)
                    {
                        LoadFileToTab(file);
                    }
                }
            }
        }

        private static bool ParseInput(string s, out double value)
        {
            // заменяем запятую на точку и пытаемся распознать
            return double.TryParse(
                       s.Replace(',', '.'),
                       NumberStyles.Float,
                       CultureInfo.InvariantCulture,
                       out value);
        }

        private void btnCalculateEps_Click(object sender, EventArgs e)
        {
            if (!ParseInput(txtThickness.Text, out double Dmm) ||
                    !ParseInput(txtDiameter.Text, out double DiameterMm))
            {
                MessageBox.Show("Введите корректные значения толщины и диаметра.",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double D = Dmm / 1000.0; // м
            double R = DiameterMm / 2000.0; // радиус в метрах
            double S = Math.PI * R * R;
            const double e0 = 8.854187817e-12;

            foreach (TabPage page in tabControlFiles.TabPages)
            {
                if (page.Controls.Count == 0 || !(page.Controls[0] is DataGridView grid))
                    continue;

                // Добавление колонок при необходимости
                if (!grid.Columns.Contains("ε"))
                {
                    grid.Columns.Add(new DataGridViewTextBoxColumn
                    {
                        Name = "ε",
                        HeaderText = "ε",
                        Width = 90,
                        ReadOnly = true
                    });
                }

                if (!grid.Columns.Contains("ε′"))
                {
                    grid.Columns.Add(new DataGridViewTextBoxColumn
                    {
                        Name = "ε′",
                        HeaderText = "ε′",
                        Width = 90,
                        ReadOnly = true
                    });
                }

                if (!grid.Columns.Contains("ε″"))
                {
                    grid.Columns.Add(new DataGridViewTextBoxColumn
                    {
                        Name = "ε″",
                        HeaderText = "ε″",
                        Width = 90,
                        ReadOnly = true
                    });
                }

                int capCol = FindColumn(grid, new[] { "C[F]", "Ёмкость C" });
                int tgCol = FindColumn(grid, TgHeaders);
                int epsCol = grid.Columns["ε"].Index;
                int epsPCol = grid.Columns["ε′"].Index;
                int epsPPCol = grid.Columns["ε″"].Index;

                if (capCol == -1 || tgCol == -1)
                {
                    MessageBox.Show($"В таблице «{page.Text}» не найдены нужные колонки C или tg δ.",
                                    "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }

                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;

                    string cStr = row.Cells[capCol].Value?.ToString()?.Replace(',', '.');
                    string tgStr = row.Cells[tgCol].Value?.ToString()?.Replace(',', '.');

                    bool okC = double.TryParse(cStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double C);
                    bool okTg = double.TryParse(tgStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double tg);

                    if (!okC)
                    {
                        row.Cells[epsCol].Value = "Ошибка";
                        row.Cells[epsPCol].Value = "Ошибка";
                        row.Cells[epsPPCol].Value = "Ошибка";
                        continue;
                    }

                    double eps = (C * D) / (e0 * S);
                    row.Cells[epsCol].Value = eps.ToString("G6", CultureInfo.InvariantCulture);

                    if (!okTg)
                    {
                        row.Cells[epsPCol].Value = "Ошибка";
                        row.Cells[epsPPCol].Value = "Ошибка";
                        continue;
                    }

                    double epsP = eps / Math.Sqrt(1 + tg * tg);
                    row.Cells[epsPCol].Value = epsP.ToString("G6", CultureInfo.InvariantCulture);

                    double epsPP = epsP * tg;
                    row.Cells[epsPPCol].Value = epsPP.ToString("G6", CultureInfo.InvariantCulture);
                }
            }

            MessageBox.Show("ε, ε′ и ε″ добавлены в таблицы.",
                            "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Берёт среднюю часть между двумя "___"; если такой схемы нет — возвращает само имя
        private static string MakeTabTitle(string fileNameWithoutExt)
        {
            var parts = fileNameWithoutExt.Split(new[] { "___" }, StringSplitOptions.None);
            return (parts.Length >= 3) ? parts[1] : fileNameWithoutExt;
        }

        // ───── небольшая утилита для поиска колонки по заголовкам
        private static int FindColumn(DataGridView grid, IEnumerable<string> headers)
        {
            foreach (var h in headers)
            {
                var col = grid.Columns
                              .Cast<DataGridViewColumn>()
                              .FirstOrDefault(c => string.Equals(c.HeaderText.Trim(), h, StringComparison.OrdinalIgnoreCase));
                if (col != null) return col.Index;
            }
            return -1;
        }

        private void btnPlotEpsChart_Click(object sender, EventArgs e)
        {
            chartEps.Series.Clear();
            chartEps.ChartAreas.Clear();
            chartEps.ChartAreas.Add(new ChartArea("MainArea"));

            // Соберём все доступные частоты из данных
            var frequencyGroups = new Dictionary<double, List<(double temperature, double eps)>>();

            foreach (var kvp in allEpsilons)
            {
                if (!double.TryParse(kvp.Key, out double temperature))
                    continue;

                foreach (var (freq, eps) in kvp.Value)
                {
                    if (!frequencyGroups.ContainsKey(freq))
                        frequencyGroups[freq] = new List<(double, double)>();

                    frequencyGroups[freq].Add((temperature, eps));
                }
            }

            foreach (var freqGroup in frequencyGroups)
            {
                var series = new Series($"f = {freqGroup.Key:G4} ГГц")
                {
                    ChartType = SeriesChartType.Line,
                    BorderWidth = 2,
                    MarkerStyle = MarkerStyle.Circle,
                    MarkerSize = 6
                };

                foreach (var (T, eps) in freqGroup.Value.OrderBy(p => p.temperature))
                {
                    series.Points.AddXY(T, eps);
                }

                chartEps.Series.Add(series);
            }

            chartEps.ChartAreas[0].AxisX.Title = "Температура (°C)";
            chartEps.ChartAreas[0].AxisY.Title = "Диэлектрическая проницаемость ε";
            chartEps.ChartAreas[0].RecalculateAxesScale();
        }

        private void cmbDevice_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            // Перестроить все вкладки под новый прибор
            tabControlFiles.TabPages.Clear();

            foreach (var fileName in rawFiles.Keys)
            {
                ParseAndDisplay(fileName); // уже написанный метод, который использует CurrentDevice
            }
        }

        private void cmbDevice_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Очистить все вкладки
            tabControlFiles.TabPages.Clear();

            // Отобразить только подходящие по текущему прибору
            foreach (var kvp in rawFiles)
            {
                if (kvp.Value.DetectedDevice == CurrentDevice)
                {
                    ParseAndDisplay(kvp.Key);
                }
            }
        }

        private void tabControlFiles_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // Определяем, по какой вкладке кликнули
            for (int i = 0; i < tabControlFiles.TabPages.Count; i++)
            {
                Rectangle r = tabControlFiles.GetTabRect(i);
                if (r.Contains(e.Location))
                {
                    StartRenamingTab(i, r);
                    break;
                }
            }
        }

        private void StartRenamingTab(int index, Rectangle tabBounds)
        {
            // Удаляем старый TextBox, если вдруг остался
            if (renameTextBox != null)
            {
                this.Controls.Remove(renameTextBox);
                renameTextBox.Dispose();
            }

            renamingTabPage = tabControlFiles.TabPages[index];

            renameTextBox = new TextBox
            {
                Bounds = new Rectangle(
                            tabControlFiles.Left + tabBounds.X + 4,
                            tabControlFiles.Top + tabBounds.Y + 4,
                            tabBounds.Width - 8,
                            tabBounds.Height - 6),
                Text = tabControlFiles.TabPages[index].Text,
                BorderStyle = BorderStyle.FixedSingle
            };

            renameTextBox.LostFocus += RenameTextBox_LostFocus;
            renameTextBox.KeyDown += RenameTextBox_KeyDown;

            this.Controls.Add(renameTextBox);
            renameTextBox.BringToFront();
            renameTextBox.Focus();
            renameTextBox.SelectAll();
        }

        private void RenameTextBox_LostFocus(object sender, EventArgs e)
        {
            ConfirmTabRename();
        }

        private void RenameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ConfirmTabRename();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                CancelTabRename();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void ConfirmTabRename()
        {
            TextBox tb = renameTextBox;
            TabPage tab = renamingTabPage;

            if (tb == null || tab == null)
                return;

            string newKey = tb.Text.Trim();
            string oldKey = tab.Tag?.ToString();

            Debug.WriteLine($"[RENAME] oldKey={oldKey}, newKey={newKey}");

            if (!string.IsNullOrEmpty(newKey) && oldKey != null && newKey != oldKey)
            {
                Debug.WriteLine($"[RENAME] Changing name from '{oldKey}' to '{newKey}'");

                // визуально
                tab.Text = newKey;

                // словари
                RenameDictionaryKey(rawFiles, oldKey, newKey);
                RenameDictionaryKey(allMeasurements, oldKey, newKey);
                RenameDictionaryKey(allEpsilons, oldKey, newKey);

                // обновить Tag
                tab.Tag = newKey;
            }

            // очистка
            tb.LostFocus -= RenameTextBox_LostFocus;
            tb.KeyDown -= RenameTextBox_KeyDown;
            this.Controls.Remove(tb);
            tb.Dispose();
            renameTextBox = null;
            renamingTabPage = null;
        }

        private void CancelTabRename()
        {
            TextBox tb = renameTextBox;
            if (tb == null) return;

            tb.LostFocus -= RenameTextBox_LostFocus;
            tb.KeyDown -= RenameTextBox_KeyDown;
            this.Controls.Remove(tb);
            tb.Dispose();

            renameTextBox = null;
            renamingTabPage = null;
        }

        // кнопка "Построить ε(T) для всех частот"
        private void btnBuildTempDependence_Click_1(object sender, EventArgs e)
        {
            // 1. Собираем данные:  freq → [ (T, tg, ε, ε′, ε″) ]
            var freqDict = new Dictionary<double, List<(double T,
                                                       double tg,
                                                       double eps,
                                                       double epsP,
                                                       double epsPP)>>();

            foreach (TabPage tempTab in tabControlFiles.TabPages)
            {
                if (tempTab.Controls.Count == 0 || !(tempTab.Controls[0] is DataGridView grid))
                    continue;

                // ── Температура из заголовка вкладки ──
                if (!double.TryParse(tempTab.Text.Replace(',', '.'),
                                     NumberStyles.Float,
                                     CultureInfo.InvariantCulture,
                                     out double T))
                    continue;                       // если заголовок‑не‑число — пропустим

                // ── индексы нужных столбцов ──
                int fCol = FindColumn(grid, new[] { "F[Hz]", "Частота f" });
                int epsCol = grid.Columns.Contains("ε") ? grid.Columns["ε"].Index : -1;
                int epsPCol = grid.Columns.Contains("ε′") ? grid.Columns["ε′"].Index : -1;
                int epsPPCol = grid.Columns.Contains("ε″") ? grid.Columns["ε″"].Index : -1;
                int tgCol = FindColumn(grid,
                                   new[] { "D[]", "D []", "tg", "tg(delta)", "tanδ", "tan(delta)" });

                if (fCol == -1 || epsCol == -1 || tgCol == -1)
                    continue;                       // без частоты, ε или tg — бессмысленно

                // ── построчно ──
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool okF = double.TryParse(row.Cells[fCol].Value?.ToString()
                                                 ?.Replace(',', '.'),
                                                 NumberStyles.Float,
                                                 CultureInfo.InvariantCulture, out double f);

                    bool okEps = double.TryParse(row.Cells[epsCol].Value?.ToString()
                                                 ?.Replace(',', '.'),
                                                 NumberStyles.Float,
                                                 CultureInfo.InvariantCulture, out double eps);

                    bool okTg = double.TryParse(row.Cells[tgCol].Value?.ToString()
                                                ?.Replace(',', '.'),
                                                NumberStyles.Float,
                                                CultureInfo.InvariantCulture, out double tg);

                    if (!okF || !okEps || !okTg) continue;

                    double epsP = double.NaN,
                           epsPP = double.NaN;

                    if (epsPCol != -1)
                        double.TryParse(row.Cells[epsPCol].Value?.ToString()
                                        ?.Replace(',', '.'),
                                        NumberStyles.Float,
                                        CultureInfo.InvariantCulture, out epsP);

                    if (epsPPCol != -1)
                        double.TryParse(row.Cells[epsPPCol].Value?.ToString()
                                        ?.Replace(',', '.'),
                                        NumberStyles.Float,
                                        CultureInfo.InvariantCulture, out epsPP);

                    if (!freqDict.ContainsKey(f))
                        freqDict[f] = new List<(double, double, double, double, double)>();

                    freqDict[f].Add((T, tg, eps, epsP, epsPP));
                }
            }

            // 2. Удаляем прежние «частотные» вкладки
            foreach (TabPage tp in tabControlFiles.TabPages
                                       .Cast<TabPage>()
                                       .Where(t => (t.Tag as string) == "freq")
                                       .ToList())
            {
                tabControlFiles.TabPages.Remove(tp);
                tp.Dispose();
            }

            // 3. Создаём новые вкладки по частотам
            foreach (var kvp in freqDict)
            {
                double f = kvp.Key;
                string tabName = f.ToString("G4", CultureInfo.InvariantCulture) + " Hz";

                var list = kvp.Value.OrderBy(v => v.T).ToList();

                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                    ScrollBars = ScrollBars.Both
                };

                grid.Columns.Add("T", "T, °C");
                grid.Columns.Add("tg", "tg(δ)");
                grid.Columns.Add("eps", "ε");
                grid.Columns.Add("epsP", "ε′");
                grid.Columns.Add("epsPP", "ε″");

                foreach (var (T, tg, eps, epsP, epsPP) in list)
                {
                    grid.Rows.Add(
                        T.ToString("G4", CultureInfo.InvariantCulture),
                        tg.ToString("G6", CultureInfo.InvariantCulture),
                        eps.ToString("G6", CultureInfo.InvariantCulture),
                        double.IsNaN(epsP) ? "" : epsP.ToString("G6", CultureInfo.InvariantCulture),
                        double.IsNaN(epsPP) ? "" : epsPP.ToString("G6", CultureInfo.InvariantCulture));
                }

                var freqTab = new TabPage(tabName) { Tag = "freq" };
                freqTab.Controls.Add(grid);
                tabControlFiles.TabPages.Add(freqTab);
            }

            MessageBox.Show("Температурные зависимости ε, ε′, ε″ и tg(δ) построены.",
                            "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void btnExportAllTables_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Выберите папку для сохранения таблиц";
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string baseFolder = dialog.SelectedPath;

                // Имя вложенной папки (проект)
                string subFolderName = PromptFolderName("Введите название для новой папки проекта:");
                if (string.IsNullOrWhiteSpace(subFolderName))
                    return;

                string rootPath = Path.Combine(baseFolder, subFolderName);
                string tempPath = Path.Combine(rootPath, "Температура");
                string freqPath = Path.Combine(rootPath, "Частотные зависимости температуры");

                Directory.CreateDirectory(tempPath);
                Directory.CreateDirectory(freqPath);

                // 1. Собираем и сортируем все вкладки: частотные отдельно, остальные отдельно
                var tempTabs = new List<TabPage>();
                var freqTabs = new List<(TabPage tab, double freq)>();

                foreach (TabPage tab in tabControlFiles.TabPages)
                {
                    if (tab.Controls.Count == 0 || !(tab.Controls[0] is DataGridView))
                        continue;

                    if ((tab.Tag as string) == "freq")
                    {
                        if (TryExtractFrequency(tab.Text, out double freq))
                            freqTabs.Add((tab, freq));
                    }
                    else
                    {
                        tempTabs.Add(tab);
                    }
                }

                // Сортируем частотные по частоте
                freqTabs.Sort((a, b) => a.freq.CompareTo(b.freq));

                // 2. Сохраняем температурные вкладки
                foreach (TabPage tab in tempTabs)
                {
                    var grid = (DataGridView)tab.Controls[0];
                    string fileName = Path.Combine(tempPath, SanitizeFileName(tab.Text) + ".txt");
                    SaveGridToFile(grid, fileName);
                }

                // 3. Сохраняем частотные вкладки
                foreach (var (tab, freq) in freqTabs)
                {
                    var grid = (DataGridView)tab.Controls[0];
                    string niceName = FormatFrequency(freq) + ".txt";
                    string fileName = Path.Combine(freqPath, SanitizeFileName(niceName));
                    SaveGridToFile(grid, fileName);
                }


                MessageBox.Show("Все таблицы успешно сохранены в:\n" + rootPath,
                                "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveGridToFile(DataGridView grid, string path)
        {
            using (var writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                var headers = grid.Columns.Cast<DataGridViewColumn>()
                                           .Select(c => c.HeaderText ?? "");
                writer.WriteLine(string.Join("\t", headers));

                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;

                    var values = row.Cells.Cast<DataGridViewCell>()
                                          .Select(c => c.Value?.ToString() ?? "");
                    writer.WriteLine(string.Join("\t", values));
                }
            }
        }

        private string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private bool TryExtractFrequency(string tabText, out double frequency)
        {
            frequency = 0;
            // ожидаем вид вроде "27900 Hz" или "2.78E+4 Hz"
            var match = Regex.Match(tabText, @"([\dEe\+\-\.]+)\s*Hz", RegexOptions.IgnoreCase);
            if (match.Success && double.TryParse(match.Groups[1].Value.Replace(',', '.'),
                                                  NumberStyles.Float,
                                                  CultureInfo.InvariantCulture,
                                                  out frequency))
            {
                return true;
            }
            return false;
        }

        private string FormatFrequency(double hz)
        {
            if (hz < 10_000)
                return hz.ToString("0", CultureInfo.InvariantCulture) + " Hz";
            if (hz < 1_000_000)
                return (hz / 1000).ToString("0.#", CultureInfo.InvariantCulture) + " kHz";
            return (hz / 1_000_000).ToString("0.#", CultureInfo.InvariantCulture) + " MHz";
        }


        private string PromptFolderName(string message)
        {
            Form prompt = new Form()
            {
                Width = 360,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Название папки",
                StartPosition = FormStartPosition.CenterScreen
            };

            var lbl = new Label() { Left = 20, Top = 20, Text = message, Width = 300 };
            var inputBox = new TextBox() { Left = 20, Top = 50, Width = 300 };

            var okBtn = new Button() { Text = "OK", Left = 170, Width = 70, Top = 80, DialogResult = DialogResult.OK };
            var cancelBtn = new Button() { Text = "Отмена", Left = 250, Width = 70, Top = 80, DialogResult = DialogResult.Cancel };

            prompt.Controls.Add(lbl);
            prompt.Controls.Add(inputBox);
            prompt.Controls.Add(okBtn);
            prompt.Controls.Add(cancelBtn);

            prompt.AcceptButton = okBtn;
            prompt.CancelButton = cancelBtn;

            return prompt.ShowDialog() == DialogResult.OK ? inputBox.Text.Trim() : null;
        }

     


    }

    public class MeasurementRow
    {
        public double Frequency { get; set; }
        public double EnergyQ { get; set; }
        public double Tg { get; set; }
        public double Impedance { get; set; }
        public double Phase { get; set; }
        public double InductanceL { get; set; }
        public double CapacityC { get; set; }
        public double ResistanceR { get; set; }
    }
}
