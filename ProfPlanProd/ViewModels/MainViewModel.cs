using System;
using System.Collections;
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
using ClosedXML.Excel;
using ControlzEx.Standard;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.Win32;
using ProfPlanProd.Commands;
using ProfPlanProd.Models;
using ProfPlanProd.ViewModels.Base;
using static System.Windows.Forms.AxHost;
using ProfPlanProd.Views;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace ProfPlanProd.ViewModels
{
    internal class MainViewModel : ViewModel
    {
        public MainViewModel() { }

        private string filePath = "";
        private string tempFilePath = "";
        private int Number = 1;
        private DataTableCollection tableCollection;
        /// <summary>
        /// Вкладка Файл 
        /// </summary>
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
            else if (table.TableName.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
            {
                ProcessAdditionalTable(table, list);
            }
            for (int i = 0; i<list.Count; i++)
            {
                list[i].PropertyChanged +=SelectedItemPropertyChanged;
            }

            TablesCollections.Add(new TableCollection(tabname, list));
            TablesCollections.SortTablesCollection();
        }

        private void ProcessAdditionalTable(DataTable table, ObservableCollection<ExcelData> list)
        {
            for (int i = 1; i < table.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(table.Rows[i][0].ToString()))
                    list.Add(new ExcelAdditional(
                        table.Rows[i][0].ToString(),
                        table.Rows[i][1].ToString(),
                        table.Rows[i][2].ToNullable<double>()
                        ));
            }
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

        #region AddTable
        private RelayCommand _addDataCommand;

        public ICommand AddDataCommand
        {
            get { return _addDataCommand ?? (_addDataCommand = new RelayCommand(AddData)); }
        }


        private void AddData(object parameter)
        {
            try
            {
                tempFilePath = GetExcelFilePath();
                if (!string.IsNullOrEmpty(tempFilePath))
                {
                    tableCollection = ReadExcelData(tempFilePath).Tables;

                    if (tableCollection.Count == 1)
                    {
                        if (_selectedComboBoxIndex == 0)
                            tableCollection[0].TableName = "П_ПИиИС";
                        else if (_selectedComboBoxIndex == 1)
                            tableCollection[0].TableName = "Ф_ПИиИС";
                        foreach (DataTable table in tableCollection)
                        {
                            DataTableInsert(table);
                        }
                        OnPropertyChanged(nameof(TablesCollections));
                        UpdateListBoxItemsSource();
                    }
                    else
                    {
                        MessageBox.Show("Можно добавить лишь 1 таблицу!");
                    }
                }
            }
            catch
            {

            }
        }

        private void DataTableInsert(DataTable table)
        {
            int ind = -1;
            if (_selectedComboBoxIndex == 0)
                ind = TablesCollections.GetTableIndexByName("П_ПИиИС", _selectedComboBoxIndex);
            else if (_selectedComboBoxIndex == 1)
                ind = TablesCollections.GetTableIndexByName("Ф_ПИиИС", _selectedComboBoxIndex);
            if (ind == -1)
            {
                Number = 1;
            }
            else
            {
                Number = TablesCollections.GetTablesCollection()[ind].ExcelDataList.Count() + 1;
            }
            string tabname = table.TableName;
            ObservableCollection<ExcelData> list = new ObservableCollection<ExcelData>();
            int rowIndex = -1;
            bool haveTeacher = false;
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
            bool exitOuterLoop = false;
            int endstring = -1;
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
                for (int i = rowIndex; i < table.Rows.Count; i++)
                {
                    if (table.Rows[i][0].ToString() == "")
                    {
                        endstring = i;
                        break;
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

            if (endstring == -1) { endstring = table.Rows.Count; }
            for (int i = rowIndex + 1; i < endstring; i++)
            {
                try
                {
                    if (haveTeacher && !string.IsNullOrWhiteSpace(table.Rows[i][0].ToString()))
                    {
                        list.Add(new ExcelModel(
                                               Number,
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
                                               Number,
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

            for (int i = 0; i<list.Count; i++)
            {
                list[i].PropertyChanged +=SelectedItemPropertyChanged;
            }
            TablesCollections.AddInOldTabCol(new TableCollection(tabname, list));
        }
        #endregion

        #region Settings and Interaction
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
                DisplayedTables =  TablesCollections.GetTablesCollectionWithP();
            }
            else if (SelectedComboBoxIndex == 1)
            {
                DisplayedTables =  TablesCollections.GetTablesCollectionWithF();
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
        #endregion

        #region Exit
        private RelayCommand _exitCommand;

        public ICommand ExitCommand
        {
            get { return _exitCommand ?? (_exitCommand = new RelayCommand(ExitFromApp)); }
        }


        private void ExitFromApp(object parameter)
        {
            Application.Current.Dispatcher.Invoke(() => Application.Current.Shutdown());
        }
        #endregion

        #region Create BaseTable
        private RelayCommand _createBaseTableCommand;

        public ICommand CreateBaseTableCommand
        {
            get { return _createBaseTableCommand ?? (_createBaseTableCommand = new RelayCommand(CreateBaseTable)); }
        }
        private void CreateBaseTable(object parameter)
        {
            try
            {
                CreateTableCollection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private async Task CreateTableCollection()
        {
            await Task.Run(() =>
            {
                if ( TablesCollections.GetTableByName("ПИиИС", SelectedComboBoxIndex) == false)
                {
                    if(SelectedComboBoxIndex == 0)
                        TablesCollections.Add(new TableCollection() { Tablename = "П_ПИиИС" });
                    else if (SelectedComboBoxIndex == 1)
                        TablesCollections.Add(new TableCollection() { Tablename = "Ф_ПИиИС" });
                }
                UpdateListBoxItemsSource();
            });

        }
        #endregion

        #region Save
        private RelayCommand _saveDataToExcelAs;
        private RelayCommand _saveDataToExcel;

        public ICommand SaveDataAsCommand
        {
            get { return _saveDataToExcelAs ?? (_saveDataToExcelAs = new RelayCommand(SaveToExcelAs)); }
        }

        public ICommand SaveDataCommand
        {
            get { return _saveDataToExcel ?? (_saveDataToExcel = new RelayCommand(SaveToExcel)); }
        }

        private void SaveToExcel(object parameter)
        {
            if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) == true || filePath == "")
                SaveToExcelAsAsync();
            else
                SaveToExcelAsync();
        }
        private async Task SaveToExcelAsync()
        {
            await Task.Run(() =>
            SaveToExcels(TablesCollections.GetTablesCollection()));
        }

        private void SaveToExcelAs(object parameter)
        {
            SaveToExcelAsAsync();
        }

        private async Task SaveToExcelAsAsync()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"Расчет Нагрузки {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                filePath = saveFileDialog.FileName;

            }
            else
            {
                return;
            }
            SaveToExcelAsync();
        }

        private async Task SaveToExcels(ObservableCollection<TableCollection> tablesCollection)
        {
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook())
                {
                    foreach (var table in tablesCollection)
                    {
                        int rowNumberAutumn = table.ExcelDataList.Count + 6;
                        int rowNumberSpring = table.ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1).Count() + table.ExcelDataList.Count() + 13;
                        var worksheet = CreateWorksheet(workbook, table);
                        PopulateWorksheet(worksheet, table);
                        if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && worksheet.Name.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) == -1 && worksheet.Name.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            worksheet.Range(1, 2, 1, 3).Merge();
                            worksheet.Cell(1, 2).Value = worksheet.Cell(3, 2).Value;
                            worksheet.Cell(1, 2).Style.Font.SetFontSize(14);
                            worksheet.Cell(1, 2).Style.Font.SetBold(true);

                            worksheet.Cell(1, 5).Value = "Всего";
                            worksheet.Cell(1, 5).Style.Font.SetFontSize(14);
                            worksheet.Cell(1, 5).Style.Font.SetBold(true);
                            //
                            worksheet.Cell(rowNumberAutumn, 2).Value = "Осень";
                            worksheet.Cell(rowNumberAutumn, 2).Style.Font.SetFontSize(14);
                            worksheet.Cell(rowNumberAutumn, 2).Style.Font.SetBold(true);

                            worksheet.Cell(rowNumberSpring, 2).Value = "Весна";
                            worksheet.Cell(rowNumberSpring, 2).Style.Font.SetFontSize(14);
                            worksheet.Cell(rowNumberSpring, 2).Style.Font.SetBold(true);
                        }
                    }
                    int frow = 2;
                    List<string> newPropertyNames = new List<string>
                {
                    "№", "Преподаватель", "Дисциплина", "Семестр(четный или нечетный)", "Группа", "Институт", "Число групп", "Подгруппа", "Форма обучения", "Число студентов", "Из них коммерч.", "Недель", "Форма отчетности", "Лекции", "Практики", "Лабораторные", "Консультации", "Зачеты", "Экзамены", "Курсовые работы", "Курсовые проекты", "ГЭК+ПриемГЭК, прием ГАК",
                    "Диплом", "РГЗ_Реф, нормоконтроль", "ПрактикаРабота, реценз диплом", "Прочее", "Всего", "Бюджетные", "Коммерческие"
                };
                    List<string> newPropertyTotalNames = new List<string>
                {
                    "ФИО", "Ставка", "Ставка(%)", "Всего", "Осень", "Весна", "Разница"
                };
                    List<string> newPropertyAdditionalNames = new List<string>
                {
                    "ФИО", "Вид работы", "Часов"
                };
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            for (int i = 0; i < newPropertyTotalNames.Count; i++)
                            {
                                worksheet.Cell(frow - 1, i + 1).Value = newPropertyTotalNames[i];
                            }
                        }
                        else if (worksheet.Name.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            for (int i = 0; i < newPropertyAdditionalNames.Count; i++)
                            {
                                worksheet.Cell(frow - 1, i + 1).Value = newPropertyAdditionalNames[i];
                            }
                        }
                        else
                        {

                            int rowNumberAutumn = TablesCollections.GetTablesCollection()[TablesCollections.GetTableIndexByName(worksheet.Name)].ExcelDataList.Count + 7;
                            int rowNumberSpring = TablesCollections.GetTablesCollection()[TablesCollections.GetTableIndexByName(worksheet.Name)].ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1).Count() + rowNumberAutumn + 7;


                            for (int i = 0; i < newPropertyNames.Count; i++)
                            {
                                worksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
                                worksheet.Cell(frow, i + 1).Style.Alignment.SetTextRotation(90);
                                if (newPropertyNames[i] != "Преподаватель")
                                    worksheet.Cell(frow, i + 1).Style.Alignment.WrapText = true;

                                worksheet.Cell(rowNumberAutumn, i + 1).Value = newPropertyNames[i];
                                worksheet.Cell(rowNumberAutumn, i + 1).Style.Alignment.SetTextRotation(90);
                                if (newPropertyNames[i] != "Преподаватель")
                                worksheet.Cell(rowNumberAutumn, i + 1).Style.Alignment.WrapText = true;

                                worksheet.Cell(rowNumberSpring, i + 1).Value = newPropertyNames[i];
                                worksheet.Cell(rowNumberSpring, i + 1).Style.Alignment.SetTextRotation(90);
                                if (newPropertyNames[i] != "Преподаватель")
                                worksheet.Cell(rowNumberSpring, i + 1).Style.Alignment.WrapText = true;
                            }
                        }
                    }
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1)
                        {

                            worksheet.Column(3).AdjustToContents(4, 4);
                            worksheet.Column(2).AdjustToContents(4, 4);
                            worksheet.Row(2).AdjustToContents(20, 20);
                            //worksheet.Rows().AdjustToContents();
                        }
                    }
                    SaveWorkbook(workbook);
                }
            });
        }

        private IXLWorksheet CreateWorksheet(XLWorkbook workbook, TableCollection table)
        {
            var worksheet = workbook.Worksheets.Add(table.Tablename);

            if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
            {
                CreateTotalHeaders(worksheet);
            }
            else if (table.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
            {
                CreateModelHeaders(worksheet);
                CreateModelHeaders(worksheet, table.ExcelDataList.Count + 7);
                CreateModelHeaders(worksheet, table.ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1).Count() + table.ExcelDataList.Count() + 14);
            }
            else
            {
                CreateAdditionalHeaders(worksheet);
            }

            return worksheet;
        }

        private void CreateAdditionalHeaders(IXLWorksheet worksheet)
        {
            int columnNumber = 1;
            foreach (var propertyInfo in typeof(ExcelAdditional).GetProperties())
            {
                worksheet.Cell(1, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Value = propertyInfo.Name;
                columnNumber++;
            }
        }

        private void CreateTotalHeaders(IXLWorksheet worksheet)
        {
            int columnNumber = 1;
            foreach (var propertyInfo in typeof(ExcelTotal).GetProperties())
            {
                worksheet.Cell(1, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, columnNumber).Value = propertyInfo.Name;
                columnNumber++;
            }
        }

        private void CreateModelHeaders(IXLWorksheet worksheet, int rowNumber = 2)
        {
            int columnNumber = 1;
            foreach (var propertyInfo in typeof(ExcelModel).GetProperties())
            {
                worksheet.Cell(rowNumber, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(rowNumber, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(rowNumber, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(rowNumber, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                if (propertyInfo.Name != "Teachers")
                {
                    worksheet.Cell(rowNumber, columnNumber).Value = propertyInfo.Name;
                    columnNumber++;
                }
            }
        }

        private void PopulateWorksheet(IXLWorksheet worksheet, TableCollection table)
        {
            int rowNumber = (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1 || table.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1) ? 2 : 3;
            int columnNumber = 1;
            int indprop = 27;
            foreach (var data in table.ExcelDataList)
            {
                if(table.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    var value1 = data.GetType().GetProperty("TeacherA")?.GetValue(data, null);
                    var value2 = data.GetType().GetProperty("TypeOfWork")?.GetValue(data, null);
                    if (value1 == null && value2 == null)
                        continue;
                }
                if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    var value1 = data.GetType().GetProperty("Teacher")?.GetValue(data, null);
                    if (value1 == null)
                        continue;
                }
                foreach (var propertyName in GetPropertyNames(data))
                {
                    worksheet.Cell(rowNumber, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(rowNumber, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(rowNumber, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(rowNumber, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    var value = data.GetType().GetProperty(propertyName)?.GetValue(data, null);
                    if(value != null && value != "")
                    {
                        if (int.TryParse(value.ToString(), out int val))
                            worksheet.Cell(rowNumber, columnNumber).Value = val;
                        else if (double.TryParse(value.ToString(), out double vald))
                            worksheet.Cell(rowNumber, columnNumber).Value = vald;
                        else
                            worksheet.Cell(rowNumber, columnNumber).Value = value.ToString();
                    }
                    else
                    {
                        worksheet.Cell(rowNumber, columnNumber).Value ="";
                    }
                    columnNumber++;
                }

                rowNumber++;
                columnNumber = 1;
            }
           

            if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && table.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
            {
                double sum = 0;
                foreach (ExcelModel excelModel in table.ExcelDataList)
                {
                    sum+=excelModel.Total.ToNullable<double>() ?? 0;
                }
                worksheet.Cell(rowNumber, indprop - 2).Value = "Итого";
                worksheet.Cell(rowNumber, indprop).Value = sum;
                worksheet.Cell(rowNumber, indprop - 2).Style.Font.SetBold(true);
                worksheet.Cell(rowNumber, indprop).Style.Font.SetBold(true);

                rowNumber = table.ExcelDataList.Count() + 8;
               int  rowNumberCh = table.ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1).Count() + table.ExcelDataList.Count() + 15;
                foreach (var data in table.ExcelDataList)
                {
                    if (data.GetTermValue()!="unnull" && (data.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        foreach (var propertyName in GetPropertyNames(data))
                        {
                            worksheet.Cell(rowNumber, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumber, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumber, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumber, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            var value = data.GetType().GetProperty(propertyName)?.GetValue(data, null);
                            if (value != null)
                            {
                                if (int.TryParse(value.ToString(), out int val))
                                    worksheet.Cell(rowNumber, columnNumber).Value = val;
                                else if (double.TryParse(value.ToString(), out double vald))
                                    worksheet.Cell(rowNumber, columnNumber).Value = vald;
                                else
                                    worksheet.Cell(rowNumber, columnNumber).Value = value.ToString();
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, columnNumber).Value ="";
                            }
                            columnNumber++;
                        }
                        rowNumber++;
                        columnNumber = 1;
                    }
                    else
                    {
                        foreach (var propertyName in GetPropertyNames(data))
                        {
                            worksheet.Cell(rowNumberCh, columnNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumberCh, columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumberCh, columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(rowNumberCh, columnNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            var value = data.GetType().GetProperty(propertyName)?.GetValue(data, null);
                            if (value != null)
                            {
                                if (int.TryParse(value.ToString(), out int val))
                                    worksheet.Cell(rowNumberCh, columnNumber).Value = val;
                                else if (double.TryParse(value.ToString(), out double vald))
                                    worksheet.Cell(rowNumberCh, columnNumber).Value = vald;
                                else
                                    worksheet.Cell(rowNumberCh, columnNumber).Value = value.ToString();
                            }
                            else
                            {
                                worksheet.Cell(rowNumberCh, columnNumber).Value ="";
                            }
                            columnNumber++;
                        }
                        rowNumberCh++;
                        columnNumber = 1;
                    }
                }
                sum=0;
                foreach (ExcelModel excelModel in table.ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1))
                {
                    sum+=excelModel.Total.ToNullable<double>() ?? 0;
                }
                worksheet.Cell(rowNumber, indprop - 2).Value = "Итого";
                worksheet.Cell(rowNumber, indprop).Value = sum;
                worksheet.Cell(rowNumber, indprop - 2).Style.Font.SetBold(true);
                worksheet.Cell(rowNumber, indprop).Style.Font.SetBold(true);
                sum=0;
                foreach (ExcelModel excelModel in table.ExcelDataList.Where(tc => (tc.GetTermValue() ?? "").IndexOf("нечет", StringComparison.OrdinalIgnoreCase) == -1))
                {
                    sum+=excelModel.Total.ToNullable<double>() ?? 0;
                }
                worksheet.Cell(rowNumberCh, indprop - 2).Value = "Итого";
                worksheet.Cell(rowNumberCh, indprop).Value = sum;
                worksheet.Cell(rowNumberCh, indprop - 2).Style.Font.SetBold(true);
                worksheet.Cell(rowNumberCh, indprop).Style.Font.SetBold(true);
            }

        }

        private IEnumerable<string> GetPropertyNames(object data)
        {
            return data is ExcelModel model
        ? typeof(ExcelModel).GetProperties().Where(p => p.Name != "Teachers").Select(p => p.Name)
        : data is ExcelAdditional additional
        ? typeof(ExcelAdditional).GetProperties().Select(p => p.Name)
        : typeof(ExcelTotal).GetProperties().Select(p => p.Name);
        }

        private void SaveWorkbook(XLWorkbook workbook)
        {
            workbook.SaveAs(filePath);
        }
        #endregion

        /// <summary>
        /// Вкладка Таблица
        /// </summary>

        #region Clear table
        private RelayCommand _clearTableCommand;

        public ICommand ClearTableCommand
        {
            get { return _clearTableCommand ?? (_clearTableCommand = new RelayCommand(ClearTable)); }
        }
        private void ClearTable(object parameter)
        {
            try
            {
                if (MessageBox.Show($"Вы уверены, что хотите очистить таблицу {SelectedTable.Tablename}?", "Очистка таблицы", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    if (SelectedTable != null && SelectedComboBoxIndex != -1)
                    {
                        TablesCollections.RemoveTableAtIndex(TablesCollections.GetTableIndexByName(SelectedTable.Tablename, SelectedComboBoxIndex));
                        UpdateListBoxItemsSource();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion


        #region Move Teachers from Plan to Fact
        private RelayCommand _moveTeachersCommand;

        public ICommand MoveTeachersCommand
        {
            get { return _moveTeachersCommand ?? (_moveTeachersCommand = new RelayCommand(MoveTeachers)); }
        }

        private void MoveTeachers(object parameter)
        {
            try
            {
                int ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
                int stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
                if (ftableindex != -1 && stableindex != -1)
                {
                    try
                    {
                        if (TablesCollections.GetTablesCollection()[stableindex].ExcelDataList.Count != 0)
                        {
                            for (int i = 0; i < TablesCollections.GetTablesCollection()[stableindex].ExcelDataList.Count; i++)
                            {
                                if (TablesCollections.GetTablesCollection()[stableindex].ExcelDataList[i] is ExcelModel excelModel && excelModel.Teacher == "")
                                {
                                    ExcelModel stableData = TablesCollections.GetTablesCollection()[stableindex].ExcelDataList[i] as ExcelModel;
                                    ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelDataList[i] as ExcelModel;

                                    if (stableData != null && ftableData != null &&
                                        stableData.Term == ftableData.Term &&
                                        stableData.Group == ftableData.Group &&
                                        stableData.Institute == ftableData.Institute &&
                                        stableData.FormOfStudy == ftableData.FormOfStudy &&
                                        ftableData.Teacher != "")
                                    {
                                        stableData.Teacher = ftableData.Teacher;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Лист Факт пустой! Поэтому данные с листа План были скопированы");
                            CreateTableCollectionsForMove();
                            ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
                            stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
                            for (int i = 0; i < TablesCollections.GetTablesCollection()[ftableindex].ExcelDataList.Count; i++)
                            {
                                ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelDataList[i] as ExcelModel;
                                TablesCollections.AddByIndex(stableindex, ftableData);
                            }
                        }

                    }
                    catch
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Лист Факт пустой! Поэтому данные с листа План были скопированы");
                    CreateTableCollectionsForMove();
                    ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
                    stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
                    for (int i = 0; i < TablesCollections.GetTablesCollection()[ftableindex].ExcelDataList.Count; i++)
                    {
                        ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelDataList[i] as ExcelModel;
                        TablesCollections.AddByIndex(stableindex, ftableData);
                    }
                }
                if (SelectedComboBoxIndex==0 && TablesCollections.GetTablesCollectionWithP().Count()>0)
                    SelectedTable = TablesCollections.GetTablesCollectionWithP()[0];
                else if (SelectedComboBoxIndex==1 && TablesCollections.GetTablesCollectionWithF().Count()>0)
                    SelectedTable = TablesCollections.GetTablesCollectionWithF()[0];
                else
                    SelectedTable = null;
                UpdateListBoxItemsSource();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void CreateTableCollectionsForMove()
        {
            if (TablesCollections.GetTableByName("П_ПИиИС", 0) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "П_ПИиИС" });
            }
            if (TablesCollections.GetTableByName("Ф_ПИиИС", 1) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "Ф_ПИиИС" });
            }
            UpdateListBoxItemsSource();
        }
        #endregion


        #region AddRow
        private RelayCommand _addRow;

        public ICommand AddRowCommand
        {
            get { return _addRow ?? (_addRow = new RelayCommand(AddRow)); }
        }
        private void AddRow(object parameter)
        {
            if(SelectedTable != null)
            {
                if (SelectedTable.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    SelectedTable.ExcelDataList.Add(new ExcelTotal());
                }
                else if (SelectedTable.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    SelectedTable.ExcelDataList.Add(new ExcelAdditional());
                }
                else
                {
                    MessageBox.Show("В данную таблицу нельзя добавить запись!");
                }
            }
        }
        #endregion

        

        #region Generate Teachers lists
        private RelayCommand _generateTeachersLists;

        public ICommand GenerateTeachersListsCommand
        {
            get { return _generateTeachersLists ?? (_generateTeachersLists = new RelayCommand(GenerateTeacher)); }
        }

        private void GenerateTeacher(object parameter)
        {
            try
            {
                GenerateTeacherAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private async Task GenerateTeacherAsync()
        {
            await Task.Run(() =>
            {
                if (SelectedComboBoxIndex != -1 && TablesCollections.GetTableIndexForGenerate("ПИиИС", SelectedComboBoxIndex) != -1)
                {
                    string prefix;
                    if (SelectedComboBoxIndex == 0)
                    {
                        prefix = "П_";
                    }
                    else
                    {
                        prefix = "Ф_";
                    }
                    int mainList = TablesCollections.GetTableIndexByName(prefix + "ПИиИС", SelectedComboBoxIndex);
                    var uniqueTeachers = TablesCollections.GetTablesCollection()[mainList].ExcelDataList
                   .Where(data => data is ExcelModel) // Фильтрация по типу ExcelModel
                   .Select(data => ((ExcelModel)data).Teacher) // Приведение к ExcelModel и выбор Teacher
                   .Distinct()
                   .ToList();
                    ObservableCollection<ExcelData> totallist = new ObservableCollection<ExcelData>();
                    foreach (var teacher in uniqueTeachers)
                    {
                        var teacherTableCollection = new TableCollection() { };

                        if (teacher.ToString() != "")
                            teacherTableCollection = new TableCollection(prefix+teacher.ToString().Split(' ')[0]);
                        else
                            teacherTableCollection = new TableCollection(prefix+"Незаполненные");
                        var teacherRows = TablesCollections.GetTablesCollection()[mainList].ExcelDataList
                        .Where(data => data is ExcelModel && ((ExcelModel)data).Teacher == teacher)
                        .ToList();
                        foreach (ExcelModel techrow in teacherRows)
                        {
                            techrow.PropertyChanged += teacherTableCollection.ExcelModel_PropertyChanged;
                            teacherTableCollection.ExcelDataList.Add(techrow);

                        }
                        teacherTableCollection.SubscribeToExcelDataChanges();
                        TablesCollections.Add(teacherTableCollection);

                        //Реализация листа Итого:

                        double? bet = null;
                        string lname, fname, mname;
                        if (teacherTableCollection.Tablename != prefix + "Незаполненные")
                        {
                            foreach (Teacher teach in TeachersManager.GetTeachers())
                            {
                                lname=teach.LastName;
                                fname=teach.FirstName;
                                mname=teach.MiddleName;
                                if ($"{lname} {fname[0]}.{mname[0]}." == teacher)
                                {
                                    bet = teach.Workload;
                                }

                            }
                            totallist.Add(new ExcelTotal(
                            teacher.IndexOf(' ') != -1 ? teacher.Substring(0, teacher.IndexOf(' ')) : teacher,
                                bet,
                            null,
                                teacherTableCollection.TotalHours,
                               teacherTableCollection.AutumnHours,
                               teacherTableCollection.SpringHours,
                                null)
                                );

                        }
                    }
                    string tabname = prefix + "Итого";
                    foreach (ExcelTotal list in totallist)
                    {
                        list.DifferenceCalc();
                    }
                    TablesCollections.Add(new TableCollection(tabname, totallist));
                    TablesCollections.Add(new TableCollection("Доп", new ObservableCollection<ExcelData>() { new ExcelAdditional() }));
                    TablesCollections.SortTablesCollection();
                    UpdateListBoxItemsSource();
                }
            });
        }
        #endregion

        /// <summary>
        /// Вкладка Преподаватели
        /// </summary>

        #region Show Teachers Window
        private RelayCommand _showTeachersWindowCommand;

        public ICommand ShowTeachersWindowCommand
        {
            get { return _showTeachersWindowCommand ?? (_showTeachersWindowCommand = new RelayCommand(ShowTeachersWindow)); }
        }

        private void ShowTeachersWindow(object obj)
        {
            try
            {
                var techerswindow = obj as Window;

                TeachersWindow teacherlist = new TeachersWindow();
                teacherlist.Owner = techerswindow;
                teacherlist.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                teacherlist.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion

        
        #region CalcReport
        private RelayCommand _loadCalcReportPlan;
        private ReportViewModel loadCalcVM = new ReportViewModel();
        public ICommand LoadCalcReportPlanCommand
        {
            get { return _loadCalcReportPlan ?? (_loadCalcReportPlan = new RelayCommand(CreateLoadCalcReportPlan)); }
        }

        private void CreateLoadCalcReportPlan(object obj)
        {
            try
            {
                _=loadCalcVM.CreateLoadCalcAsync(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private RelayCommand _loadCalcReportFact;
        public ICommand LoadCalcReportFactCommand
        {
            get { return _loadCalcReportFact ?? (_loadCalcReportFact = new RelayCommand(CreateLoadCalcReportFact)); }
        }

        private void CreateLoadCalcReportFact(object obj)
        {
            try
            {
                _=loadCalcVM.CreateLoadCalcAsync(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion


        #region Individual Plan report
        private RelayCommand _individualPlanReport;

        public ICommand LoadIndividualPlanReportCommmand
        {
            get { return _individualPlanReport ?? (_individualPlanReport = new RelayCommand(CreateIndividualPlanReport)); }
        }

        private void CreateIndividualPlanReport(object obj)
        {
            try
            {
                loadCalcVM.CreateIndividualPlan(SelectedComboBoxIndex);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion
    }
}
