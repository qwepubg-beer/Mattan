using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Mattan
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataGrid transportGrid;
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
        }
        public void BuildTransportGrid(int suppliersCount, int consumersCount)
        {
            transportTable = new DataTable();

            // 1. Первый столбец для названий поставщиков
            transportTable.Columns.Add("Поставщик \\ Потребитель", typeof(string));

            // 2. Столбцы для потребителей (B1, B2, ...)
            for (int j = 0; j < consumersCount; j++)
            {
                transportTable.Columns.Add($"B{j + 1}", typeof(double));
            }

            // 3. Последний столбец для запасов
            transportTable.Columns.Add("Запасы", typeof(double));

            // 4. Заполняем строки поставщиков (A1, A2, ...)
            for (int i = 0; i < suppliersCount; i++)
            {
                DataRow row = transportTable.NewRow();
                row[0] = $"A{i + 1}";

                // Заполняем нулями по умолчанию
                for (int j = 1; j <= consumersCount + 1; j++)
                {
                    row[j] = 0;
                }
                transportTable.Rows.Add(row);
            }

            // 5. Добавляем последнюю строку для потребностей
            DataRow demandRow = transportTable.NewRow();
            demandRow[0] = "Потребности";
            for (int j = 1; j <= consumersCount; j++)
            {
                demandRow[j] = 0;
            }
            // Нижняя правая ячейка (пересечение запасов и потребностей) остается пустой (DBNull)
            transportTable.Rows.Add(demandRow);

            // Привязываем таблицу к DataGrid
            transportGrid.ItemsSource = transportTable.DefaultView;

            // Запрещаем редактировать первый столбец с названиями (A1, A2, Потребности)
            transportGrid.AutoGeneratingColumn += (s, e) =>
            {
                if (e.Column.Header.ToString() == "Поставщик \\ Потребитель")
                {
                    e.Column.IsReadOnly = true;
                }
            };
        }
        public void ReadTransportData(out double[,] costs, out double[] supply, out double[] demand)
        {
            int suppliersCount = transportTable.Rows.Count - 1;
            int consumersCount = transportTable.Columns.Count - 2;

            costs = new double[suppliersCount, consumersCount];
            supply = new double[suppliersCount];
            demand = new double[consumersCount];

            // Считываем тарифы и запасы
            for (int i = 0; i < suppliersCount; i++)
            {
                DataRow row = transportTable.Rows[i];

                for (int j = 0; j < consumersCount; j++)
                {
                    // +1, так как 0-й столбец — это название поставщика
                    costs[i, j] = Convert.ToDouble(row[j + 1]);
                }

                // Последний столбец — запасы
                supply[i] = Convert.ToDouble(row[consumersCount + 1]);
            }

            // Считываем потребности из последней строки
            DataRow demandRow = transportTable.Rows[suppliersCount];
            for (int j = 0; j < consumersCount; j++)
            {
                demand[j] = Convert.ToDouble(demandRow[j + 1]);
            }
        }

        private void next_Click(object sender, RoutedEventArgs e)
        {
            if(RowCb.SelectedIndex!=-1 && ColumnCb.SelectedIndex != -1)
            {
                int countRow = RowCb.SelectedIndex + 1;
                int countColumn = ColumnCb.SelectedIndex + 1;
                BuildTransportGrid(countRow, countColumn);
            }

        }
    }
}
