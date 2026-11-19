using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Mattan
{
    public partial class MainWindow : Window
    {
        private DataGrid transportGrid;      // исходная таблица (тарифы, запасы, потребности)
        private DataGrid solutionGrid;       // таблица решения
        private DataTable transportTable;

        public MainWindow()
        {
            InitializeComponent();

            transportGrid = new DataGrid
            {
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserSortColumns = false,
                AutoGenerateColumns = true,
                Margin = new Thickness(10)
            };
            Grid.SetRow(transportGrid, 1);
            Main.Children.Add(transportGrid);

            solutionGrid = new DataGrid
            {
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserSortColumns = false,
                AutoGenerateColumns = true,
                IsReadOnly = true,
                Margin = new Thickness(10),
                Background = new SolidColorBrush(Color.FromRgb(245, 250, 245))
            };
            Grid.SetRow(solutionGrid, 1);
            // Чтобы решения не перекрывались, положим их в один Grid с отступом сверху
            // Проще: создадим отдельный Grid поверх transportGrid
            Panel.SetZIndex(solutionGrid, 10);
            solutionGrid.Visibility = Visibility.Collapsed;
            Main.Children.Add(solutionGrid);
        }

        // ============================================================
        //  1. Построение пустой таблицы
        // ============================================================
        public void BuildTransportGrid(int suppliersCount, int consumersCount)
        {
            transportTable = new DataTable();

            transportTable.Columns.Add("Поставщик \\ Потребитель", typeof(string));

            for (int j = 0; j < consumersCount; j++)
                transportTable.Columns.Add($"B{j + 1}", typeof(double));

            transportTable.Columns.Add("Запасы", typeof(double));

            for (int i = 0; i < suppliersCount; i++)
            {
                DataRow row = transportTable.NewRow();
                row[0] = $"A{i + 1}";
                for (int j = 1; j <= consumersCount + 1; j++)
                    row[j] = 0.0;
                transportTable.Rows.Add(row);
            }

            DataRow demandRow = transportTable.NewRow();
            demandRow[0] = "Потребности";
            for (int j = 1; j <= consumersCount; j++)
                demandRow[j] = 0.0;
            transportTable.Rows.Add(demandRow);

            transportGrid.ItemsSource = transportTable.DefaultView;
            solutionGrid.Visibility = Visibility.Collapsed;

            transportGrid.AutoGeneratingColumn += (s, e) =>
            {
                if (e.Column.Header.ToString() == "Поставщик \\ Потребитель")
                    e.Column.IsReadOnly = true;
            };
        }

        // ============================================================
        //  2. Чтение данных из грида
        // ============================================================
        public void ReadTransportData(out double[,] costs, out double[] supply, out double[] demand)
        {
            int suppliersCount = transportTable.Rows.Count - 1;
            int consumersCount = transportTable.Columns.Count - 2;

            costs = new double[suppliersCount, consumersCount];
            supply = new double[suppliersCount];
            demand = new double[consumersCount];

            for (int i = 0; i < suppliersCount; i++)
            {
                DataRow row = transportTable.Rows[i];
                for (int j = 0; j < consumersCount; j++)
                    costs[i, j] = Convert.ToDouble(row[j + 1]);
                supply[i] = Convert.ToDouble(row[consumersCount + 1]);
            }

            DataRow demandRow = transportTable.Rows[suppliersCount];
            for (int j = 0; j < consumersCount; j++)
                demand[j] = Convert.ToDouble(demandRow[j + 1]);
        }

        // ============================================================
        //  3. Кнопка «Построить»
        // ============================================================
        private void next_Click(object sender, RoutedEventArgs e)
        {
            if (RowCb.SelectedIndex != -1 && ColumnCb.SelectedIndex != -1)
            {
                int countRow = RowCb.SelectedIndex + 1;
                int countColumn = ColumnCb.SelectedIndex + 1;
                BuildTransportGrid(countRow, countColumn);
            }
            else
            {
                MessageBox.Show("Выберите количество поставщиков и потребителей.",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ============================================================
        //  4. Кнопка «Решить»
        // ============================================================
        private void Answer_Click(object sender, RoutedEventArgs e)
        {
            if (transportTable == null)
            {
                MessageBox.Show("Сначала постройте таблицу.", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (Choose.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите метод построения опорного плана.", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                ReadTransportData(out double[,] costs, out double[] supply, out double[] demand);

                // Проверка баланса
                double sumSupply = 0, sumDemand = 0;
                foreach (var s in supply) sumSupply += s;
                foreach (var d in demand) sumDemand += d;

                if (Math.Abs(sumSupply - sumDemand) > 1e-6)
                {
                    MessageBox.Show(
                        $"Задача не закрытая!\nСумма запасов = {sumSupply}\nСумма потребностей = {sumDemand}\n" +
                        "Добавьте фиктивного поставщика/потребителя.",
                        "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                double[,] plan = Choose.SelectedIndex == 0
                    ? NorthWestCorner(costs, supply, demand)
                    : MinCostMethod(costs, supply, demand);
                double totalCost = 0;
                for (int i = 0; i < plan.GetLength(0); i++)
                    for (int j = 0; j < plan.GetLength(1); j++)
                        totalCost += plan[i, j] * costs[i, j];
                ShowSolution(plan, costs, supply, demand, totalCost);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при решении: " + ex.Message, "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ============================================================
        //  5. Северо-западный угол
        // ============================================================
        private static double[,] NorthWestCorner(double[,] costs, double[] supply, double[] demand)
        {
            int n = supply.Length, m = demand.Length;
            double[,] plan = new double[n, m];
            double[] s = (double[])supply.Clone();
            double[] d = (double[])demand.Clone();

            int i = 0, j = 0;
            while (i < n && j < m)
            {
                double x = Math.Min(s[i], d[j]);
                plan[i, j] = x;
                s[i] -= x;
                d[j] -= x;

                if (s[i] < 1e-9) i++;
                if (d[j] < 1e-9) j++;
            }
            return plan;
        }

        // ============================================================
        //  6. Метод наименьшего элемента
        // ============================================================
        private static double[,] MinCostMethod(double[,] costs, double[] supply, double[] demand)
        {
            int n = supply.Length, m = demand.Length;
            double[,] plan = new double[n, m];
            double[] s = (double[])supply.Clone();
            double[] d = (double[])demand.Clone();
            bool[,] used = new bool[n, m];

            while (true)
            {
                // Ищем минимальную стоимость среди незакрытых ячеек
                double minCost = double.MaxValue;
                int bi = -1, bj = -1;

                for (int i = 0; i < n; i++)
                {
                    if (s[i] < 1e-9) continue;
                    for (int j = 0; j < m; j++)
                    {
                        if (d[j] < 1e-9) continue;
                        if (used[i, j]) continue;
                        if (costs[i, j] < minCost)
                        {
                            minCost = costs[i, j];
                            bi = i; bj = j;
                        }
                    }
                }

                if (bi == -1) break; // всё распределено

                double x = Math.Min(s[bi], d[bj]);
                plan[bi, bj] = x;
                s[bi] -= x;
                d[bj] -= x;

                if (s[bi] < 1e-9)
                    for (int j = 0; j < m; j++) used[bi, j] = true;
                if (d[bj] < 1e-9)
                    for (int i = 0; i < n; i++) used[i, bj] = true;

                // Если обнулились обе — закрываем ячейку, но не сдвигаем обе координаты
                if (s[bi] < 1e-9 && d[bj] < 1e-9)
                    used[bi, bj] = true;
            }
            return plan;
        }
        private void ShowSolution(double[,] plan, double[,] costs,
                                  double[] supply, double[] demand, double totalCost)
        {
            int n = supply.Length, m = demand.Length;
            DataTable solTable = new DataTable();
            solTable.Columns.Add("Поставщик \\ Потребитель", typeof(string));
            for (int j = 0; j < m; j++)
            solTable.Columns.Add($"B{j + 1} (c={costs[0, j]:0.##})", typeof(string));
            solTable.Columns.Add("Запасы", typeof(double));
            for (int i = 0; i < n; i++)
            {
                DataRow row = solTable.NewRow();
                row[0] = $"A{i + 1}";
                for (int j = 0; j < m; j++)
                {
                    row[j + 1] = plan[i, j] > 1e-9
                        ? $"{plan[i, j]:0.##}"
                        : "-";
                }
                row[m + 1] = supply[i];
                solTable.Rows.Add(row);
            }
            DataRow demandRow = solTable.NewRow();
            demandRow[0] = "Потребности";
            for (int j = 0; j < m; j++)
            {
                demandRow[j + 1] = demand[j].ToString("0.##");
            }
            demandRow[m+1] = totalCost.ToString("0.##");
            solTable.Rows.Add(demandRow);
            solutionGrid.ItemsSource = solTable.DefaultView;
            solutionGrid.Visibility = Visibility.Visible;
            MessageBox.Show($"Общая стоимость перевозок: {totalCost:0.##}",
                            "Решение найдено", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "Выберите файл с транспортной задачей"
            };
            if (dlg.ShowDialog() != true) return;

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range used = null;

            try
            {
                app = new Excel.Application { Visible = false };
                wb = app.Workbooks.Open(dlg.FileName);
                ws = (Excel.Worksheet)wb.Sheets[1];

                used = ws.UsedRange;

                int firstRow = used.Row;               // первая строка UsedRange
                int firstCol = used.Column;            // первый столбец UsedRange
                int lastRow = firstRow + used.Rows.Count - 1;
                int lastCol = firstCol + used.Columns.Count - 1;

                // --- Ищем последний столбец по заголовку "Запасы" ---
                int colSupply = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    object h = used.Cells[firstRow, c].Value2;
                    if (h != null && h.ToString().Trim().StartsWith("Запас", StringComparison.OrdinalIgnoreCase))
                    {
                        colSupply = c;
                        break;
                    }
                }
                if (colSupply == -1) throw new Exception("Не найден столбец «Запасы».");

                // --- Ищем последнюю строку по заголовку "Потреб..." ---
                int rowDemand = -1;
                for (int r = firstRow; r <= lastRow; r++)
                {
                    object h = used.Cells[r, firstCol].Value2;
                    if (h != null && h.ToString().Trim().StartsWith("Потреб", StringComparison.OrdinalIgnoreCase))
                    {
                        rowDemand = r;
                        break;
                    }
                }
                if (rowDemand == -1) throw new Exception("Не найдена строка «Потребности».");

                // --- Размеры ---
                // Данные поставщиков: со строки (firstRow + 1) до (rowDemand - 1)
                // Данные потребителей: со столбца (firstCol + 1) до (colSupply - 1)
                int suppliers = rowDemand - (firstRow + 1);
                int consumers = colSupply - (firstCol + 1);

                if (suppliers < 1 || consumers < 1)
                    throw new Exception("Не удалось определить размеры таблицы. Проверьте заголовки.");

                BuildTransportGrid(suppliers, consumers);

                // --- Читаем тарифы и запасы ---
                for (int i = 0; i < suppliers; i++)
                {
                    DataRow row = transportTable.Rows[i];
                    int excelRow = firstRow + 1 + i;

                    for (int j = 0; j < consumers; j++)
                    {
                        int excelCol = firstCol + 1 + j;
                        object v = used.Cells[excelRow, excelCol].Value2;
                        row[j + 1] = TryToDouble(v);
                    }

                    object s = used.Cells[excelRow, colSupply].Value2;
                    row[consumers + 1] = TryToDouble(s);
                }

                // --- Читаем потребности ---
                DataRow demandRow = transportTable.Rows[suppliers];
                for (int j = 0; j < consumers; j++)
                {
                    int excelCol = firstCol + 1 + j;
                    object v = used.Cells[rowDemand, excelCol].Value2;
                    demandRow[j + 1] = TryToDouble(v);
                }

                transportGrid.ItemsSource = transportTable.DefaultView;
                solutionGrid.Visibility = Visibility.Collapsed;

                MessageBox.Show($"Импортировано: {suppliers} поставщиков × {consumers} потребителей.",
                                "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка импорта Excel: " + ex.Message, "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (used != null) Marshal.ReleaseComObject(used);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Безопасное преобразование ячейки Excel в double
        private static double TryToDouble(object value)
        {
            if (value == null) return 0.0;
            if (value is double d) return d;
            if (double.TryParse(value.ToString().Replace(',', '.'),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double res))
                return res;
            return 0.0;
        }
    }
}