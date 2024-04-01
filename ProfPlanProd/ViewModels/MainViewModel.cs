using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelDataReader;
using Microsoft.Win32;
using ProfPlanProd.Commands;
using ProfPlanProd.Models;
using ProfPlanProd.ViewModels.Base;

namespace ProfPlanProd.ViewModels
{
    internal class MainViewModel : ViewModel
    {
        public MainViewModel() { }

        private string filePath = "";
        private int Number = 1;
        private DataTableCollection tableCollection;
        #region OpenFileCommand
        private RelayCommand _loadDataCommand;

        public ICommand LoadDataCommand
        {
            get { return _loadDataCommand ?? (_loadDataCommand = new RelayCommand(LoadDataToTablesCollection)); }
        }


        private void LoadDataToTablesCollection(object parameter)
        {
            try
            {
                filePath = GetExcelFilePath();
                if (!string.IsNullOrEmpty(filePath))
                {
                     tableCollection = ReadExcelData(filePath).Tables;
                TablesCollections.Clear();
                foreach (DataTable table in tableCollection)
                {
                    ProcessDataTable(table);
                }
                }
                UpdateListBoxItemsSource();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error:{ex.Message}");
            }
        }

        private string GetExcelFilePath()
        {
            var openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx" };

            return openFileDialog.ShowDialog() == true ? openFileDialog.FileName : null;
        }


        private DataSet ReadExcelData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                    });
                }
            }
        }

        private void ProcessDataTable(DataTable table)
        {
            string tabname = table.TableName;
            ObservableCollection<ExcelData> list = new ObservableCollection<ExcelData>();
            int rowIndex = -1;
            bool haveTeacher = false;
            bool exitOuterLoop = false;
            int endstring = -1;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count - 1; j++)
                {
                    if (table.Rows[i][j].ToString().Trim() == "Дисциплина")
                    {
                        rowIndex = i;
                        break;
                    }
                }
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count - 1; j++)
                {
                    if (table.Rows[i][j].ToString().Trim() == "Дисциплина")
                    {
                        rowIndex = i;

                        exitOuterLoop = true;
                        break;
                    }
                }

                if (exitOuterLoop)
                {
                    break;
                }
            }
            if (rowIndex != -1)
            {
                for (int i = rowIndex; i < table.Rows.Count; i++)
                {
                    if (table.Rows[i][0].ToString() == "")
                    {
                        endstring = i;
                        break;
                    }
                }
            }

            for (int j = 0; j < table.Columns.Count - 1; j++)
            {
                if (rowIndex != -1 && table.Rows[rowIndex][j].ToString().Trim() == "Преподаватель")
                {
                    haveTeacher = true;
                    break;
                }
            }

            if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 &&
                                    table.TableName.IndexOf("доп", StringComparison.OrdinalIgnoreCase) == -1)
            {
                if (endstring == -1)
                {
                    endstring = table.Rows.Count;
                }

                for (int i = rowIndex + 1; i < endstring; i++)
                {
                    try
                    {
                        if (haveTeacher && !string.IsNullOrWhiteSpace(table.Rows[i][0].ToString()))
                        {
                            list.Add(new ExcelModel(
                                                   Convert.ToInt32(table.Rows[i][0].ToString()),
                                                   table.Rows[i][1].ToString(),
                                                   table.Rows[i][2].ToString(),
                                                   table.Rows[i][3].ToString(),
                                                   table.Rows[i][4].ToString(),
                                                   table.Rows[i][5].ToString(),
                                                   table.Rows[i][6].ToNullable<int>(),
                                                   table.Rows[i][7].ToString(),
                                                   table.Rows[i][8].ToString(),
                                                   table.Rows[i][9].ToNullable<int>(),
                                                   table.Rows[i][10].ToNullable<int>(),
                                                   table.Rows[i][11].ToNullable<int>(),
                                                   table.Rows[i][12].ToString(),
                                                   table.Rows[i][13].ToNullable<int>(),
                                                   table.Rows[i][14].ToNullable<double>(),
                                                   table.Rows[i][15].ToNullable<double>(),
                                                   table.Rows[i][16].ToNullable<double>(),
                                                   table.Rows[i][17].ToNullable<double>(),
                                                   table.Rows[i][18].ToNullable<double>(),
                                                   table.Rows[i][19].ToNullable<double>(),
                                                   table.Rows[i][20].ToNullable<double>(),
                                                   table.Rows[i][21].ToNullable<double>(),
                                                   table.Rows[i][22].ToNullable<double>(),
                                                   table.Rows[i][23].ToNullable<double>(),
                                                   table.Rows[i][24].ToNullable<double>(),
                                                   table.Rows[i][25].ToNullable<double>(),
                                                   table.Rows[i][26].ToNullable<double>(),
                                                   table.Rows[i][27].ToNullable<double>(),
                                                   table.Rows[i][28].ToNullable<double>()));
                            Number++;
                        }
                        else if (!haveTeacher)
                        {
                            list.Add(new ExcelModel(
                                                   Convert.ToInt32(table.Rows[i][0].ToString()),
                                                   "",
                                                   table.Rows[i][1].ToString(),
                                                   table.Rows[i][2].ToString(),
                                                   table.Rows[i][3].ToString(),
                                                   table.Rows[i][4].ToString(),
                                                   table.Rows[i][5].ToNullable<int>(),
                                                   table.Rows[i][6].ToString(),
                                                   table.Rows[i][7].ToString(),
                                                   table.Rows[i][8].ToNullable<int>(),
                                                   table.Rows[i][9].ToNullable<int>(),
                                                   table.Rows[i][10].ToNullable<int>(),
                                                   table.Rows[i][11].ToString(),
                                                   table.Rows[i][12].ToNullable<double>(),
                                                   table.Rows[i][13].ToNullable<double>(),
                                                   table.Rows[i][14].ToNullable<double>(),
                                                   table.Rows[i][15].ToNullable<double>(),
                                                   table.Rows[i][16].ToNullable<double>(),
                                                   table.Rows[i][17].ToNullable<double>(),
                                                   table.Rows[i][18].ToNullable<double>(),
                                                   table.Rows[i][19].ToNullable<double>(),
                                                   table.Rows[i][20].ToNullable<double>(),
                                                   table.Rows[i][21].ToNullable<double>(),
                                                   table.Rows[i][22].ToNullable<double>(),
                                                   table.Rows[i][23].ToNullable<double>(),
                                                   table.Rows[i][24].ToNullable<double>(),
                                                   table.Rows[i][25].ToNullable<double>(),
                                                   table.Rows[i][26].ToNullable<double>(),
                                                   table.Rows[i][27].ToNullable<double>()));
                            Number++;
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show($"Error adding data: {ex.Message}");
                    }
                }
            }
            else if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
            {
                ProcessTotalTable(table, list);
            }
            //else if (table.TableName.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
            //{
            //    ProcessAdditionalTable(table, list);
            //}
            for (int i = 0; i<list.Count; i++)
            {
                list[i].PropertyChanged +=SelectedItemPropertyChanged;
            }

            TablesCollections.Add(new TableCollection(tabname, list));
        }

        private void ProcessTotalTable(DataTable table, ObservableCollection<ExcelData> list)
        {
            bool hasBetPer = false;

            for (int i = 1; i < table.Columns.Count; i++)
            {
                if (table.Rows[0][i].ToString().IndexOf("%", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    hasBetPer = true;
                    break;
                }
            }

            if (hasBetPer != true)
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(table.Rows[i][0].ToString()))
                        list.Add(new ExcelTotal(
                            table.Rows[i][0].ToString(),
                            table.Rows[i][1].ToNullable<int>(),
                            null,
                            table.Rows[i][2].ToNullable<double>(),
                            table.Rows[i][3].ToNullable<double>(),
                            table.Rows[i][4].ToNullable<double>(),
                            Math.Round(Convert.ToDouble(table.Rows[i][5].ToNullable<double>()), 2)
                            ));
                }
            else
            {
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(table.Rows[i][0].ToString()))
                        list.Add(new ExcelTotal(
                            table.Rows[i][0].ToString(),
                             table.Rows[i][1].ToNullable<int>(),
                            table.Rows[i][2].ToNullable<double>(),
                            table.Rows[i][3].ToNullable<double>(),
                            table.Rows[i][4].ToNullable<double>(),
                            table.Rows[i][5].ToNullable<double>(),
                            table.Rows[i][6].ToNullable<double>()
                            ));
                }
            }
        }

        private void SelectedItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Teacher")
            {
                var changedItem = (ExcelModel)sender;
                var newTeacher = changedItem.Teacher;

                foreach (ExcelModel item in _selectedItems)
                {
                    if (item != changedItem)
                    {
                        item.Teacher = newTeacher;
                    }
                }
            }
        }

        #endregion


        private int _selectedComboBoxIndex;
        private ObservableCollection<TableCollection> _displayedTables;
        private TableCollection _selectedTable;
        private ObservableCollection<ExcelData> _selectedItems = new ObservableCollection<ExcelData>();
        public ObservableCollection<ExcelData> SelectedItems
        { get { return _selectedItems; } }

        public int SelectedComboBoxIndex
        {
            get { return _selectedComboBoxIndex; }
            set
            {
                if (_selectedComboBoxIndex != value)
                {
                    _selectedComboBoxIndex = value;
                    OnPropertyChanged(nameof(SelectedComboBoxIndex));

                    // Обновляем ItemsSource для ListBox в зависимости от выбранного элемента в ComboBox
                    UpdateListBoxItemsSource();
                }
            }
        }

        public ObservableCollection<TableCollection> DisplayedTables
        {
            get { return _displayedTables; }
            set
            {
                if (_displayedTables != value)
                {
                    _displayedTables = value;
                    OnPropertyChanged(nameof(DisplayedTables));
                }
            }
        }

        private void UpdateListBoxItemsSource()
        {
            if (SelectedComboBoxIndex == 0)
            {
                DisplayedTables =  TablesCollections.GetTablesCollection();
            }
        }

        public TableCollection SelectedTable
        {
            get { return _selectedTable; }
            set
            {
                if (_selectedTable != value)
                {
                    _selectedTable = value;
                    OnPropertyChanged(nameof(SelectedTable));
                }
            }
        }

        private Dock _tabStripPlacement = Dock.Left;
        public Dock TabStripPlacement
        {
            get { return _tabStripPlacement; }
            set
            {
                if (_tabStripPlacement != value)
                {
                    _tabStripPlacement = value;
                    OnPropertyChanged(nameof(TabStripPlacement));
                }
            }
        }

        private string _placementIcon = "ArrowLeft";
        public string PlacementIcon
        {
            get { return _placementIcon; }
            set
            {
                if (_placementIcon != value)
                {
                    _placementIcon = value;
                    OnPropertyChanged(nameof(PlacementIcon));
                }
            }
        }

        private RelayCommand _selectTabItemsPlacement;
        public ICommand SelectTabItemsPlacementCommand
        {
            get { return _selectTabItemsPlacement ?? (_selectTabItemsPlacement = new RelayCommand(SelectTabItemsPlacement)); }
        }
        private void SelectTabItemsPlacement(object parameter)
        {
            SelectTabItemsPlacementAsync();

        }
        private async Task SelectTabItemsPlacementAsync()
        {
            await Task.Run(() =>
            {
                switch (_tabStripPlacement)
                {
                    case Dock.Top:
                        _tabStripPlacement = Dock.Right;
                        PlacementIcon = "ArrowRight";
                        break;
                    case Dock.Right:
                        _tabStripPlacement = Dock.Bottom;
                        PlacementIcon = "ArrowDown";
                        break;
                    case Dock.Bottom:
                        _tabStripPlacement = Dock.Left;
                        PlacementIcon = "ArrowLeft";
                        break;
                    case Dock.Left:
                        _tabStripPlacement = Dock.Top;
                        PlacementIcon = "ArrowUp";
                        break;
                    default:
                        _tabStripPlacement = Dock.Left;
                        PlacementIcon = "ArrowLeft";
                        break;
                }
                OnPropertyChanged(nameof(TabStripPlacement));
                OnPropertyChanged(nameof(PlacementIcon));
            });
        }



    }
}
