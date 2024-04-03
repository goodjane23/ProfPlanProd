using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ProfPlanProd.Models;
using ProfPlanProd.ViewModels.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using static System.Windows.Forms.AxHost;

namespace ProfPlanProd.ViewModels
{
    internal class ReportViewModel : ViewModel
    {
        public ObservableCollection<TableCollection> TablesCollectionTeacherSumP { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherSumListP { get; set; }
        public ObservableCollection<TableCollection> TablesCollectionTeacherSumF { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherSumListF { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherForCheck { get; set; }

        /// <summary>
        /// Бланк нагрузки
        /// </summary>
        private double _state;
        public double State
        {
            get { return _state; }
            set
            {
                if (_state != value)
                {
                    _state = value;
                    OnPropertyChanged(nameof(State));
                    OnStateChanged();
                }
            }
        }
        protected virtual void OnStateChanged()
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
        public event EventHandler StateChanged;

        #region LoadCalc
        public async Task CreateLoadCalcAsync(int index)
        {
            try
            {
                bool hasTables = false;
                if(index == 0)
                        hasTables = TablesCollections.GetTableByName("П_ПИиИС", index);
                else if(index == 1)
                    hasTables = TablesCollections.GetTableByName("Ф_ПИиИС", index);
                if (hasTables)
                {
                    string directoryPath = GetSaveFilePath(index);
                    if (string.IsNullOrEmpty(directoryPath))
                        return;
                    await Task.Run(() =>
                    {
                        SumAllTeachersTables(index);
                        if (index == 0)
                            SaveToExcel(TablesCollectionTeacherSumP, directoryPath);
                        else
                            SaveToExcel(TablesCollectionTeacherSumF, directoryPath);
                    });
                }
                else
                {
                    MessageBox.Show("Ошибка! Таблицы П_/Ф_ ПИиИС отсуствуют!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private string GetSaveFilePath(int index)
        {
            string prefix;
            if (index == 0)
            {
                prefix = "_ПЛАН_";
            }
            else
            {
                prefix = "_ФАКТ_";
            }
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"Бланк{prefix}Нагрузки {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                return saveFileDialog.FileName;

            return null;
        }

        private void SumAllTeachersTables(int index)
        {

            if (index == 1)
            {
                TablesCollectionTeacherSumF = new ObservableCollection<TableCollection>();

                TablesCollectionTeacherSumListF = new ObservableCollection<TableCollection>();

            }
            else if (index == 0)
            {
                TablesCollectionTeacherSumP = new ObservableCollection<TableCollection>();

                TablesCollectionTeacherSumListP = new ObservableCollection<TableCollection>();

            }
            TablesCollectionTeacherForCheck = new ObservableCollection<TableCollection>();
            TablesCollections.SortTablesCollection();
            int completedTables = 0;
            int totalTables = TablesCollections.GetTablesCollection().Count();
            foreach (var tableCollection in TablesCollections.GetTablesCollection())
            {
                if (tableCollection.Tablename.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Незаполненные", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    double? bet = null;
                    double? betPercent = null;
                    double? totalHours = null;
                    double? autumnHours = null;
                    double? springHours = null;
                    string prefix;
                    if (index == 0)
                    {
                        prefix = "П_";
                    }
                    else
                    {
                        prefix = "Ф_";
                    }
                    var tablesCollection = index == 0 ? TablesCollections.GetTablesCollectionWithP() : TablesCollections.GetTablesCollectionWithF();

                    foreach (var tableCol in tablesCollection)
                    {
                        if (tableCol.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            foreach (ExcelTotal exRow in tableCol)
                            {
                                if (prefix + exRow.Teacher == tableCollection.Tablename)
                                {
                                    bet = exRow.Bet;
                                    betPercent = exRow.BetPercent;
                                    totalHours = exRow.TotalHours;
                                    autumnHours = exRow.AutumnHours;
                                    springHours = exRow.SpringHours;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                    TableCollection sumTableCollection;
                    TableCollection sumTableCollectionTwo;
                    //Сумма колонок для Итого
                    ExcelModel sumOdd = CalculateSum(tableCollection, "нечет");
                    ObservableCollection<ExcelModel> sumOddList = TotalSemesterCalculation(tableCollection, "нечет");

                    ExcelModel sumEven = CalculateSum(tableCollection, "чет");
                    ObservableCollection<ExcelModel> sumEvenList = TotalSemesterCalculation(tableCollection, "чет");

                    double? autumnIndex = autumnHours/totalHours;
                    double? springIndex = springHours/totalHours;
                    if (bet == null)
                        bet = totalHours;
                    if (betPercent == 1 || betPercent == null)
                    {
                        sumTableCollection = new TableCollection($"{tableCollection.Tablename}");


                        if (bet!=null)
                        {
                            ProcessBet(index, sumTableCollection, ref sumOddList, ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                            if (index == 0)
                                TablesCollectionTeacherSumP.Add(sumTableCollection);
                            else
                                TablesCollectionTeacherSumF.Add(sumTableCollection);
                        }

                    }
                    else
                    {
                        if (bet!=null)
                        {
                            if (betPercent > 1)
                            {
                                sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                                sumTableCollectionTwo = new TableCollection($"{tableCollection.Tablename} {betPercent - 1}");

                                sumTableCollection = ProcessBet(index, sumTableCollection, ref sumOddList, ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                if (index == 0)
                                    TablesCollectionTeacherSumP.Add(sumTableCollection);
                                else
                                    TablesCollectionTeacherSumF.Add(sumTableCollection);

                                sumTableCollectionTwo = ProcessBet(index, sumTableCollectionTwo, ref sumOddList, ref sumEvenList, bet*(betPercent - 1), betPercent, autumnIndex, springIndex);
                                if (index == 0)
                                    TablesCollectionTeacherSumP.Add(sumTableCollectionTwo);
                                else
                                    TablesCollectionTeacherSumF.Add(sumTableCollectionTwo);
                            }
                            else
                            {
                                sumTableCollection = new TableCollection($"{tableCollection.Tablename} {betPercent}");

                                ProcessBet(index, sumTableCollection, ref sumOddList, ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                if (index == 0)
                                    TablesCollectionTeacherSumP.Add(sumTableCollection);
                                else
                                    TablesCollectionTeacherSumF.Add(sumTableCollection);

                            }
                        }

                    }
                }
                completedTables++;
                double progress = (double)completedTables / totalTables * 60;

                State = (double)progress;
            }

        }

        private TableCollection ProcessBet(int index, TableCollection sumTableCollection, ref ObservableCollection<ExcelModel> sumOddList, ref ObservableCollection<ExcelModel> sumEvenList, double? bet, double? betPercent, double? autumnIndex = null, double? springIndex = null)
        {
            TableCollection sumOddListOneBet = new TableCollection();
            TableCollection sumEvenListOneBet = new TableCollection();

            TableCollection ListForIPPlan = new TableCollection($"{sumTableCollection.Tablename}");

            double? sum = 0, betValue, dif;

            betValue = bet * autumnIndex;

            if (betPercent<1)
            {
                betValue = bet * autumnIndex * betPercent;
            }
            var sortedList = sumOddList.OrderByDescending(x => x.SumProperties());

            foreach (ExcelModel excelModel in sortedList)
            {
                if (betValue > sum)
                {
                    dif = betValue - sum;
                    if (dif >= excelModel.SumProperties())
                    {
                        sum += excelModel.SumProperties();
                        sumOddListOneBet.ExcelDataList.Add(excelModel);
                        ListForIPPlan.ExcelDataList.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }
            DeleteItemsFromObsCol(ref sumOddList, sumOddListOneBet);

            sum = 0;
            betValue = bet * springIndex;
            if (betPercent<1)
            {
                betValue = bet * springIndex * betPercent;
            }

            sortedList = sumEvenList.OrderByDescending(x => x.SumProperties());

            foreach (ExcelModel excelModel in sortedList)
            {
                if (betValue > sum)
                {
                    dif = betValue - sum;
                    if (dif >= excelModel.SumProperties())
                    {
                        sum += excelModel.SumProperties();
                        sumEvenListOneBet.ExcelDataList.Add(excelModel);
                        ListForIPPlan.ExcelDataList.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }
            DeleteItemsFromObsCol(ref sumEvenList, sumEvenListOneBet);

            //List<string> lists = new List<string>();
            //for (int i = 0; i<sumOddListOneBet.ExcelData.Count; i++)
            //{
            //    lists.Add(($"{(sumOddListOneBet.ExcelData[i] as ExcelModel).Discipline} - {(sumOddListOneBet.ExcelData[i] as ExcelModel).Group}  - {(sumOddListOneBet.ExcelData[i] as ExcelModel).Total}"));
            //}
            //MessageBox.Show($"{sumTableCollection.Tablename}{string.Join("\n ", lists)}");
            //lists.Clear();
            //for (int i = 0; i<sumEvenListOneBet.ExcelData.Count; i++)
            //{
            //    lists.Add(($"{(sumEvenListOneBet.ExcelData[i] as ExcelModel).Discipline} - {(sumEvenListOneBet.ExcelData[i] as ExcelModel).Group}  - {(sumEvenListOneBet.ExcelData[i] as ExcelModel).Total}"));
            //}
            //MessageBox.Show($"{sumTableCollection.Tablename}{string.Join("\n ", lists)}");
            TableCollection tempTabCol = new TableCollection($"{sumTableCollection.Tablename}");
            for(int i = 0; i<sumOddListOneBet.ExcelDataList.Count; i++)
            {
                tempTabCol.ExcelDataList.Add(sumOddListOneBet.ExcelDataList[i]);
            }
            for (int j = 0; j<sumEvenListOneBet.ExcelDataList.Count; j++)
            {
                tempTabCol.ExcelDataList.Add(sumEvenListOneBet.ExcelDataList[j]);
            }
            TablesCollectionTeacherForCheck.Add(tempTabCol);
            ExcelModel sumOddOneBet = CalculateSum(sumOddListOneBet, "нечет");
            ExcelModel sumEvenOneBet = CalculateSum(sumEvenListOneBet, "чет");
            sumTableCollection.ExcelDataList.Add(sumOddOneBet);
            sumTableCollection.ExcelDataList.Add(sumEvenOneBet);
            if (index == 0)
                TablesCollectionTeacherSumListP.Add(ListForIPPlan);
            else
                TablesCollectionTeacherSumListF.Add(ListForIPPlan);
            return sumTableCollection;
        }
        private void DeleteItemsFromObsCol(ref ObservableCollection<ExcelModel> collection, TableCollection tabCol)
        {
            foreach (ExcelModel tab in tabCol)
            {
                collection.Remove(tab);
            }
        }
        private ExcelModel CalculateSum(TableCollection tableCollection, string term)
        {
            var sumModel = new ExcelModel(0, "", "", term, "", "", null, "", "", null, null, null,
                null, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            foreach (var excelModel in tableCollection.ExcelDataList.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals(term, StringComparison.OrdinalIgnoreCase)))
            {
                if (excelModel.Lectures!=null)
                    sumModel.Lectures += excelModel.Lectures;
                if (excelModel.Consultations != null)
                    sumModel.Consultations += excelModel.Consultations;
                if (excelModel.Laboratory != null)
                    sumModel.Laboratory += excelModel.Laboratory;
                if (excelModel.Practices != null)
                    sumModel.Practices += excelModel.Practices;
                if (excelModel.Tests != null)
                    sumModel.Tests += excelModel.Tests;
                if (excelModel.Exams != null)
                    sumModel.Exams += excelModel.Exams;
                if (excelModel.CourseProjects != null)
                    sumModel.CourseProjects += excelModel.CourseProjects;
                if (excelModel.CourseWorks != null)
                    sumModel.CourseWorks += excelModel.CourseWorks;
                if (excelModel.Diploma != null)
                    sumModel.Diploma += excelModel.Diploma;
                if (excelModel.RGZ != null)
                    sumModel.RGZ += excelModel.RGZ;
                if (excelModel.GEKAndGAK != null)
                    sumModel.GEKAndGAK += excelModel.GEKAndGAK;
                if (excelModel.ReviewDiploma != null)
                    sumModel.ReviewDiploma += excelModel.ReviewDiploma;
                if (excelModel.Other != null)
                    sumModel.Other += excelModel.Other;
            }
            if (sumModel.Lectures == 0)
                sumModel.Lectures = null;
            if (sumModel.Consultations == 0)
                sumModel.Consultations = null;
            if (sumModel.Laboratory == 0)
                sumModel.Laboratory = null;
            if (sumModel.Practices == 0)
                sumModel.Practices = null;
            if (sumModel.Tests == 0)
                sumModel.Tests = null;
            if (sumModel.Exams == 0)
                sumModel.Exams = null;
            if (sumModel.CourseProjects == 0)
                sumModel.CourseProjects = null;
            if (sumModel.CourseWorks == 0)
                sumModel.CourseWorks = null;
            if (sumModel.Diploma == 0)
                sumModel.Diploma = null;
            if (sumModel.RGZ == 0)
                sumModel.RGZ = null;
            if (sumModel.GEKAndGAK == 0)
                sumModel.GEKAndGAK = null;
            if (sumModel.ReviewDiploma == 0)
                sumModel.ReviewDiploma = null;
            if (sumModel.Other == 0)
                sumModel.Other = null;

            return sumModel;
        }
        private ObservableCollection<ExcelModel> TotalSemesterCalculation(TableCollection tableCollection, string term)
        {
            ObservableCollection<ExcelModel> ex = new ObservableCollection<ExcelModel>();
            foreach (var excelModel in tableCollection.ExcelDataList.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals(term, StringComparison.OrdinalIgnoreCase)))
            {
                ex.Add(excelModel);
            }
            return ex;
        }

        public void SaveToExcel(ObservableCollection<TableCollection> tablesCollection, string directoryPath)
        {
            using (var workbook = new XLWorkbook())
            {
                var fworksheet = workbook.Worksheets.Add("Бланк нагрузки");
                int frow = 3;
                // Добавление заголовков
                int columnNumber = 1;

                fworksheet.Cell(frow, columnNumber++).Value = "Teacher";

                List<string> propertyNames = new List<string>();

                foreach (var propertyInfo in typeof(ExcelModel).GetProperties())
                {
                    if (propertyInfo.Name == "Lectures" || propertyInfo.Name == "Consultations" || propertyInfo.Name == "Laboratory" || propertyInfo.Name == "Practices" || propertyInfo.Name == "Tests" || propertyInfo.Name == "Exams" || propertyInfo.Name == "CourseProjects" || propertyInfo.Name == "CourseWorks" || propertyInfo.Name == "Diploma" || propertyInfo.Name == "RGZ" || propertyInfo.Name == "GEKAndGAK" || propertyInfo.Name == "ReviewDiploma" || propertyInfo.Name == "Other")
                    {
                        fworksheet.Cell(frow, columnNumber).Value = propertyInfo.Name;
                        propertyNames.Add(propertyInfo.Name);
                        columnNumber++;
                    }
                }

                fworksheet.Cell(3, columnNumber).Value = "TotalSemester";

                // Заполнение данных - первые элементы
                int rowNumber = 4;
                int completedTables = 0;
                int totalTables = tablesCollection.Count;
                foreach (var tableCollection in tablesCollection)
                {
                    string teacherName = CreateTeachersNameForForm(tableCollection.Tablename);

                    if (tableCollection.ExcelDataList.Count >= 1)
                    {
                        var excelModel = tableCollection.ExcelDataList[0];
                        columnNumber = 1;
                        fworksheet.Cell(rowNumber, columnNumber++).Value = teacherName;

                        // Сумма колонок
                        double totalSemester = propertyNames.Sum(propertyName => Convert.ToDouble(typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null) ?? 0));
                        foreach (var propertyName in propertyNames)
                        {
                            var value = typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null);
                            fworksheet.Cell(rowNumber, columnNumber++).Value = value != null ? value.ToString() : "";
                        }

                        fworksheet.Cell(rowNumber, columnNumber).Value = totalSemester.ToString();

                        rowNumber++;
                    }
                    completedTables++;
                    double progress = (double)completedTables / totalTables * 20;

                    State = (double)progress + 60;
                }

                //// Дублирую колонки
                int headerRow = 3;

                rowNumber=4;

                //columnNumber = 1;
                int cnum = 16;
                foreach (var propertyName in propertyNames)
                {
                    fworksheet.Cell(headerRow, cnum++).Value = propertyName;
                }
                fworksheet.Cell(headerRow, cnum).Value = "TotalSemester";

                // Заполнение данных - вторые элементы
                foreach (var tableCollection in tablesCollection)
                {

                    if (tableCollection.ExcelDataList.Count == 2)
                    {
                        var excelModel = tableCollection.ExcelDataList[1];

                        cnum = 16;

                        // Сумма
                        double totalSemester = propertyNames.Sum(propertyName => Convert.ToDouble(typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null) ?? 0));
                        foreach (var propertyName in propertyNames)
                        {
                            var value = typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null);
                            fworksheet.Cell(rowNumber, cnum++).Value = value != null ? value.ToString() : "";
                        }


                        fworksheet.Cell(rowNumber, cnum).Value = totalSemester.ToString();

                        rowNumber++;
                    }
                }
                AdjustWorksheetLayout(fworksheet, frow, workbook);
                // Сохранение в файл
                SaveToExcelAdditionalLists(TablesCollectionTeacherForCheck, workbook);
                workbook.SaveAs(directoryPath);
            }
        }

        private string CreateTeachersNameForForm(string teacherName)
        {
            if (teacherName.StartsWith("П_") || teacherName.StartsWith("Ф_"))
            {
                teacherName = teacherName.Substring(2);
            }
            foreach (var teach in TeachersManager.GetTeachers())
            {
                if (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ')[0] == Regex.Replace(teach.LastName.Trim(), @"\s+", " "))
                {
                    if (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ').Length > 1)
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName + "\n" + teach.AcademicDegree + " " +
                        teach.Position +"\n" + (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ').Length > 1 ? Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ')[1] : "") + " ставки";
                    }
                    else if (teach.AcademicDegree!="" && teach.Position!="")
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName + "\n" + teach.AcademicDegree + " " + teach.Position;

                    }
                    else
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName;

                    }

                }
            }
            return teacherName;
        }

        private void AdjustWorksheetLayout(IXLWorksheet fworksheet, int frow, XLWorkbook workbook)
        {
            SwapAndInsertColumns(fworksheet, frow);

            // Задаем заголовок для нового столбца
            fworksheet.Cell(frow, 42).Value = "ИТОГО ЗА ГОД";
            // Заполняем значениями новый столбец на основе данных из других столбцов
            for (int row = 4; row <= fworksheet.RowsUsed().Count()+2; row++)
            {
                var value21 = fworksheet.Cell(row, 21).Value.ToString().ToNullable<double>();
                var value41 = fworksheet.Cell(row, 41).Value.ToString().ToNullable<double>();
                fworksheet.Cell(row, 42).Value = (value21 + value41).ToString();
            }
            frow = 1;
            //sworksheet.Range(frow, 1, frow, 3).Merge();
            fworksheet.Range(frow, 1, frow, 3).Merge();
            fworksheet.Cell(frow, 1).Value = "Первое полугодие";
            fworksheet.Cell(frow, 1).Style.Font.SetBold(true);

            fworksheet.Range(frow, 22, frow, 27).Merge();
            fworksheet.Cell(frow, 22).Value = "Второе полугодие";
            fworksheet.Cell(frow, 22).Style.Font.SetBold(true);

            fworksheet.Range(fworksheet.Cell(2, 8), fworksheet.Cell(2, 14)).Merge();
            fworksheet.Cell(2, 8).Value = "Руководство";

            fworksheet.Range(fworksheet.Cell(2, 28), fworksheet.Cell(2, 34)).Merge();
            fworksheet.Cell(2, 28).Value = "Руководство";

            List<string> newPropertyNames = GetPropertyNamesForColumns();

            for (int col = 1; col <= 43; col++)
            {
                if ((col < 8 || col > 14) && (col < 28 || col > 34)) // Проверяем, что колонка не входит в диапазоны 8-14 и 28-34
                {
                    fworksheet.Range(fworksheet.Cell(2, col), fworksheet.Cell(3, col)).Merge();
                }
            }
            for (int i = 1; i < newPropertyNames.Count; i++)
            {
                if (i<7 ||i>13)
                {
                    fworksheet.Cell(2, i + 1 + 20).Value = newPropertyNames[i];
                    fworksheet.Cell(2, i + 1).Value = newPropertyNames[i];
                }

            }
            SetStyleForWorksheet(fworksheet, workbook);
        }
        private void SetStyleForWorksheet(IXLWorksheet fworksheet, XLWorkbook workbook)
        {
            fworksheet.Cell(2, 42).Value = "ИТОГО ЗА ГОД";

            var styleArial6 = workbook.Style;
            styleArial6.Alignment.TextRotation = 90;
            styleArial6.Alignment.WrapText = true;
            for (int row = 2; row <= fworksheet.RowsUsed().Count()+1; row++)
            {
                for (int col = 2; col <= fworksheet.ColumnsUsed().Count(); col++)
                {
                    fworksheet.Cell(row, col).Style = styleArial6;
                    if (fworksheet.Cell(row, col).Value.ToString() != null)
                    {
                        if (int.TryParse(fworksheet.Cell(row, col).Value.ToString(), out int val))
                            fworksheet.Cell(row, col).Value = val;
                        else if (double.TryParse(fworksheet.Cell(row, col).Value.ToString(), out double vald))
                        {
                            fworksheet.Cell(row, col).Value = vald;
                            fworksheet.Cell(row, col).Style.NumberFormat.Format = "0.##";
                        }
                        else
                            fworksheet.Cell(row, col).Value = fworksheet.Cell(row, col).Value.ToString();
                    }
                    else
                    {
                        fworksheet.Cell(row, col).Value ="";
                    }

                }
            }
            for (int row = 2; row <= fworksheet.RowsUsed().Count(); row++)
            {
                for (int col = 1; col <= fworksheet.ColumnsUsed().Count(); col++)
                {
                    fworksheet.Cell(row, col).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    fworksheet.Cell(row, col).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    fworksheet.Cell(row, col).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    fworksheet.Cell(row, col).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
            }

            fworksheet.Column(1).Style.Alignment.WrapText = true;
            fworksheet.Cell(2, 8).Style = workbook.Style.Alignment.SetTextRotation(0);
            fworksheet.Cell(2, 28).Style = workbook.Style.Alignment.SetTextRotation(0);
            fworksheet.Columns().AdjustToContents(2);
            fworksheet.Rows(2, 3).AdjustToContents();

        }
        private List<string> GetPropertyNamesForColumns()
        {
            return new List<string>
                {
                    "Преподаватель", "Чтение лекций", "Консультации", "Лабораторные работы",
                    "Практические занятия", "Зачеты", "Экзамены", "Курсовыми проектами",
                    "Курсовыми работами", "Дипломными работами", "Учебной практикой", "Произв. практикой",
                    "УИРС", "Аспирантами и соискат.", "РГР", "Консультации для заочников",
                    "Рецензирование контр. Работ заочников", "ГЭК",
                    "Проверка контрольных работ", "Другие виды работ", "ИТОГО ЗА СЕМЕСТР"
                };
        }

        private void SwapAndInsertColumns(IXLWorksheet fworksheet, int frow)
        {
            SwapColumns(fworksheet, 3, 5);
            SwapColumns(fworksheet, 8, 9);
            SwapColumns(fworksheet, 10, 11);
            SwapColumns(fworksheet, 11, 12);
            SwapColumns(fworksheet, 17, 19);
            SwapColumns(fworksheet, 22, 23);
            SwapColumns(fworksheet, 24, 25);
            SwapColumns(fworksheet, 25, 26);

            fworksheet.Column(11).InsertColumnsBefore(4);
            fworksheet.Column(16).InsertColumnsBefore(2);
            fworksheet.Column(31).InsertColumnsBefore(4);
            fworksheet.Column(36).InsertColumnsBefore(2);
            List<string> newPropertyNames = GetPropertyNamesForColumns();


            fworksheet.Cell(frow, 1).Value="";
            for (int i = 1; i < newPropertyNames.Count; i++)
            {
                fworksheet.Cell(frow, i + 1 + 20).Value = newPropertyNames[i];
                fworksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
            }
            if (fworksheet.Column(42) == null)
            {
                fworksheet.Column(41).InsertColumnsAfter(1);
            }
        }
        public static void SwapColumns(IXLWorksheet worksheet, int column1Index, int column2Index)
        {
            int startRow = worksheet.FirstRowUsed().RowNumber();
            int endRow = worksheet.LastRowUsed().RowNumber();

            for (int row = startRow; row <= endRow; row++)
            {
                var tempValue = worksheet.Cell(row, column1Index).Value;
                worksheet.Cell(row, column1Index).Value = worksheet.Cell(row, column2Index).Value;
                worksheet.Cell(row, column2Index).Value = tempValue;
            }
        }


        //addSave
        private  void SaveToExcelAdditionalLists(ObservableCollection<TableCollection> tablesCollection, XLWorkbook workbook)
        {
            int completedTables = 0;
            int totalTables = tablesCollection.Count;
                    foreach (var table in tablesCollection)
                    {
                        table.SortExcelDataListByNumber();
                        var worksheet = CreateWorksheet(workbook, table);
                        PopulateWorksheet(worksheet, table);
                        if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && worksheet.Name.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) == -1 && worksheet.Name.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
                        {
                    worksheet.Range(1, 2, 1, 3).Merge();
                    worksheet.Cell(1, 2).Value = worksheet.Cell(3, 2).Value;
                            worksheet.Cell(1, 2).Style.Font.SetFontSize(14);
                            worksheet.Cell(1, 2).Style.Font.SetBold(true);
                        }
                completedTables++;
                double progress = (double)completedTables / totalTables * 20;

                State = (double)progress + 80;
            }
                    int frow = 2;
                    List<string> newPropertyNames = new List<string>
                {
                    "№", "Преподаватель", "Дисциплина", "Семестр(четный или нечетный)", "Группа", "Институт", "Число групп", "Подгруппа", "Форма обучения", "Число студентов", "Из них коммерч.", "Недель", "Форма отчетности", "Лекции", "Практики", "Лабораторные", "Консультации", "Зачеты", "Экзамены", "Курсовые работы", "Курсовые проекты", "ГЭК+ПриемГЭК, прием ГАК",
                    "Диплом", "РГЗ_Реф, нормоконтроль", "ПрактикаРабота, реценз диплом", "Прочее", "Всего", "Бюджетные", "Коммерческие"
                };
                    foreach (var worksheet in workbook.Worksheets)
                    {
                            if (worksheet.Name.IndexOf("Бланк нагрузки", StringComparison.OrdinalIgnoreCase) == -1)
                        {

                            for (int i = 0; i < newPropertyNames.Count; i++)
                            {
                                worksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
                                worksheet.Cell(frow, i + 1).Style.Alignment.SetTextRotation(90);
                                if (newPropertyNames[i] != "Преподаватель")
                                    worksheet.Cell(frow, i + 1).Style.Alignment.WrapText = true;
                            }
                        }
                    }
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1)
                        {
                    //worksheet.Column(2).AdjustToContents(2,2);
                    worksheet.Column(3).AdjustToContents(4, 4);
                    worksheet.Column(2).AdjustToContents(4, 4);
                    worksheet.Row(2).AdjustToContents(20,20);
                }
                    }
        }

        private IXLWorksheet CreateWorksheet(XLWorkbook workbook, TableCollection table)
        {
            var worksheet = workbook.Worksheets.Add(table.Tablename);

                CreateModelHeaders(worksheet);            

            return worksheet;
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
        #endregion

        private ObservableCollection<List<IndividualPlan>> IPListP = new ObservableCollection<List<IndividualPlan>>(), IPListF = new ObservableCollection<List<IndividualPlan>>();

        public async Task CreateIndividualPlan(int index)
        {
            try
            {
                var workbook = new XLWorkbook();

                string directoryPath = GetSaveFilePathForIP();
                if (string.IsNullOrEmpty(directoryPath))
                    return;
                await Task.Run(() =>
                {

                    bool hasTablesP = false, hasTablesF = false;
                    hasTablesP = TablesCollections.GetTableByName("П_ПИиИС", 0);
                    hasTablesF = TablesCollections.GetTableByName("Ф_ПИиИС", 1);

                    if(hasTablesP && hasTablesF)
                    {
                        SumAllTeachersTables(0);
                        SumAllTeachersTables(1);
                    }
                    else if(hasTablesF)
                    {
                        MessageBox.Show("Таблица П_ПИиИС отсуствует! П_ПИиИС не будет учитываться при составлении отчета.");
                        SumAllTeachersTables(1);
                    }
                    else if (hasTablesP)
                    {
                        MessageBox.Show("Таблица Ф_ПИиИС отсуствует! Расчеты будут проводится по П_ПИиИС.");
                        SumAllTeachersTables(0);
                    }
                    else
                    {
                        MessageBox.Show("Ошибка! Таблица П_/Ф_ПИиИС отсуствует!");
                    }
                    if (TablesCollectionTeacherSumListP!= null && TablesCollectionTeacherSumListP.Count>0)
                        foreach (TableCollection tab in TablesCollectionTeacherSumListP)
                        {
                            IPListP.Add(CreateIPList(tab, workbook));
                        }
                    if (TablesCollectionTeacherSumListF!=null && TablesCollectionTeacherSumListF.Count>0)
                        foreach (TableCollection tab in TablesCollectionTeacherSumListF)
                        {
                            IPListF.Add(CreateIPList(tab, workbook));
                        }
                    ObservableCollection<TableCollection> SomeTab;
                    if (IPListF != null && IPListF.Count>0)
                        SomeTab = TablesCollectionTeacherSumListF;
                    else
                        SomeTab = TablesCollectionTeacherSumListP;
                    int completedTables = 0;
                    int totalTables = SomeTab.Count;
                    for (int i = 0; i< SomeTab.Count; i++)
                    {
                        int row = 1;
                        string tabName = SomeTab[i].Tablename;
                        if (SomeTab[i].Tablename.StartsWith("П_") || SomeTab[i].Tablename.StartsWith("Ф_"))
                        {
                            tabName = SomeTab[i].Tablename.Substring(2); // Удаление "П_" или "Ф_" из начала строки
                        }
                        var worksheet = workbook.Worksheets.Add(tabName);

                        //Итого
                        worksheet.Range(row, 8, row + 1, 8).Merge();
                        worksheet.Cell(row, 8).Value = "Виды учебных занятий (работ)";
                        worksheet.Cell(row, 8).Style.Font.SetBold(true);
                        worksheet.Range(row, 9, row, 10).Merge();
                        worksheet.Cell(row, 9).Value = "нечетный семестр";
                        worksheet.Cell(row, 9).Style.Font.SetBold(true);
                        worksheet.Range(row, 11, row, 12).Merge();
                        worksheet.Cell(row, 11).Value = "четный семестр";
                        worksheet.Cell(row, 11).Style.Font.SetBold(true);
                        worksheet.Range(row, 13, row, 14).Merge();
                        worksheet.Cell(row, 13).Value = "Итого за уч.год";
                        worksheet.Cell(row, 13).Style.Font.SetBold(true);
                        row++;
                        worksheet.Cell(row, 9).Value = "План";
                        worksheet.Cell(row, 11).Value = "План";
                        worksheet.Cell(row, 13).Value = "План";
                        worksheet.Cell(row, 14).Value = "Факт";
                        worksheet.Cell(row, 10).Value = "Факт";
                        worksheet.Cell(row, 12).Value = "Факт";
                        for (int j = 8; j<15; j++)
                        {
                            worksheet.Cell(row, j).Style.Font.SetBold(true);
                        }

                        row++;
                        int r = row;

                        int col1, col2, col3;
                        if (index == 0)
                        {
                            col1 = 9;
                            col2 = 11;
                            col3 = 13;
                        }
                        else
                        {
                            col1 = 10;
                            col2 = 12;
                            col3 = 14;
                        }
                        //var groupedByTypeOfWork = IPList.GroupBy(ip => new { ip.TypeOfWork, ip.Term });
                        //List<(string TypeOfWork, string Term, double? TotalHours)> resultList = new List<(string, string, double?)>();
                        if (IPListF.Count>0 && IPListP.Count>0)
                        {
                            col1 = 9;
                            col2 = 11;
                            col3 = 13;
                            CreateTotal(IPListP[i], index, r, ref row, worksheet, col1, col2, col3);
                            col1 = 10;
                            col2 = 12;
                            col3 = 14;
                            CreateTotal(IPListF[i], index, r, ref row, worksheet, col1, col2, col3);
                            SumAllTables(worksheet, row, r);
                            for (int z = 1; z < row+1; z++)
                            {
                                for (int col = 8; col <= 14; col++)
                                {
                                    var cell = worksheet.Cell(z, col);
                                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                }
                            }
                            row+=4;

                            WorkWithWorkSheet(worksheet, IPListF[i]);
                        }
                        else if (IPListF.Count>0)
                        {
                            CreateTotal(IPListF[i], index, r, ref row, worksheet, col1, col2, col3);
                            SumTables(worksheet, row, r, col1, col2, col3);
                            for (int z = 1; z < row+1; z++)
                            {
                                for (int col = 8; col <= 14; col++)
                                {
                                    var cell = worksheet.Cell(z, col);
                                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                }
                            }
                            row+=4;

                            WorkWithWorkSheet(worksheet, IPListF[i]);
                        }
                        else
                        {
                            CreateTotal(IPListP[i], index, r, ref row, worksheet, col1, col2, col3);
                            SumTables(worksheet, row, r, col1, col2, col3);
                            for (int z = 1; z < row + 1; z++)
                            {
                                for (int col = 8; col <= 14; col++)
                                {
                                    var cell = worksheet.Cell(z, col);
                                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                }
                            }
                            row+=4;

                            WorkWithWorkSheet(worksheet, IPListP[i]);
                        }
                        completedTables++;
                        double progress = (double)completedTables / totalTables * 100;

                        State = (double)progress;
                    }
                });


                workbook.SaveAs(directoryPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetSaveFilePathForIP()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"ИП {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                return saveFileDialog.FileName;

            return null;
        }

        private List<IndividualPlan> CreateIPList(TableCollection tab, IXLWorkbook workbook)
        {
            List<IndividualPlan> IPList = new List<IndividualPlan>();
            //Перечень предметов

            foreach (ExcelModel excel in tab.ExcelDataList)
            {
                if (excel.Total != 0 && excel.Total!=null)
                {
                    IPList.Add(excel.FormulateIndividualPlan());
                    IPList[IPList.Count - 1].TypeOfWork = excel.GetTypeOfWork();
                }
            }
            var groupedPlans = IPList.GroupBy(ip => new { ip.Discipline, ip.TypeOfWork, ip.Term, ip.Group, ip.GroupCount, ip.SubGroup, ip.Branch })
                                .Select(group => new IndividualPlan(
                                    group.Key.Discipline,
                                    group.Key.TypeOfWork,
                                    group.Key.Term,
                                    group.Key.Group,
                                    group.Key.GroupCount,
                                    group.Key.SubGroup,
                                    group.Key.Branch,
                                    group.Sum(ip => ip.Hours)
                                ))
                                .ToList();

            IPList = groupedPlans;
            IPList = IPList.OrderBy(ip => ip.Discipline)
           .ThenBy(ip => ip.TypeOfWork)
           .ThenBy(ip => ip.SubGroup)
           .ThenBy(ip => ip.Group)

           .ToList();
            return IPList;
        }

        public void WorkWithWorkSheet(IXLWorksheet worksheet, List<IndividualPlan> IPList)
        {

            int row = 1;
            int r = row;

            worksheet.Cell(row, 1).Value = "Дисциплина";
            worksheet.Cell(row, 2).Value = "Вид работы";
            worksheet.Cell(row, 3).Value = "Группа";
            worksheet.Cell(row, 4).Value = "Подгруппа";
            worksheet.Cell(row, 5).Value = "Филиал";
            worksheet.Cell(row, 6).Value = "Часы";
            for (int j = 1; j<7; j++)
            {
                worksheet.Cell(row, j).Style.Font.SetBold(true);
            }
            row++;
            worksheet.Range(row, 1, row, 5).Merge();
            worksheet.Cell(row, 1).Value = "Четный семестр";

            row++;
            foreach (IndividualPlan ip in IPList)
            {
                if (ip.Term.IndexOf("нечет", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    worksheet.Cell(row, 1).Value = ip.Discipline;
                    worksheet.Cell(row, 2).Value = ip.TypeOfWork;
                    worksheet.Cell(row, 3).Value = ip.Group;
                    worksheet.Cell(row, 4).Value = ip.SubGroup;
                    worksheet.Cell(row, 5).Value = ip.Branch;
                    worksheet.Cell(row, 6).Value = ip.Hours;
                    row++;
                }
            }
            var range = worksheet.Range(3, 6, row, 6);
            double sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(2, 6).Value = sum;
            worksheet.Cell(2, 6).Style.NumberFormat.Format = "0.##";
            worksheet.Cell(2, 6).Style.Font.SetBold(true);
            for (int z = r; z < row; z++)
            {
                for (int col = 1; col <= 6; col++)
                {
                    var cell = worksheet.Cell(z, col);
                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
            }
            r=row;

            worksheet.Cell(row, 1).Value = "Дисциплина";
            worksheet.Cell(row, 2).Value = "Вид работы";
            worksheet.Cell(row, 3).Value = "Группа";
            worksheet.Cell(row, 4).Value = "Подгруппа";
            worksheet.Cell(row, 5).Value = "Филиал";
            worksheet.Cell(row, 6).Value = "Часы";
            for (int j = 1; j<7; j++)
            {
                worksheet.Cell(row, j).Style.Font.SetBold(true);
            }
            row++;
            worksheet.Range(row, 1, row, 5).Merge();
            worksheet.Cell(row, 1).Value = "Нетный семестр";
            row++;
            int srow = row;
            foreach (IndividualPlan ip in IPList)
            {
                if (ip.Term.IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    worksheet.Cell(row, 1).Value = ip.Discipline;
                    worksheet.Cell(row, 2).Value = ip.TypeOfWork;
                    worksheet.Cell(row, 3).Value = ip.Group;
                    worksheet.Cell(row, 4).Value = ip.SubGroup;
                    worksheet.Cell(row, 5).Value = ip.Branch;
                    worksheet.Cell(row, 6).Value = ip.Hours;
                    row++;
                }
            }
            range = worksheet.Range(srow, 6, row, 6);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(srow - 1, 6).Value = sum;
            worksheet.Cell(srow - 1, 6).Style.NumberFormat.Format = "0.##";
            worksheet.Cell(srow - 1, 6).Style.Font.SetBold(true);
            for (int z = r; z < row; z++)
            {
                for (int col = 1; col <= 6; col++)
                {
                    var cell = worksheet.Cell(z, col);
                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
            }
            worksheet.Column(1).Style.Alignment.WrapText = true;
            worksheet.Columns(1, 7).AdjustToContents();
            worksheet.Columns(8, 14).AdjustToContents(1);
            worksheet.Rows().AdjustToContents();
        }

        private void CreateTotal(List<IndividualPlan> IPList, int index, int r, ref int row, IXLWorksheet worksheet, int col1, int col2, int col3)
        {
            int ind;
            var evenTermList = IPList.Where(ip => ip.Term == "чет").ToList();
            var oddTermList = IPList.Where(ip => ip.Term == "нечет").ToList();

            // Группировка и подсчет суммы часов для каждого типа работы
            var evenTermGrouped = evenTermList.GroupBy(ip => ip.TypeOfWork)
                                              .Select(group => new { TypeOfWork = group.Key, TotalHours = group.Sum(ip => ip.Hours) })
                                              .ToList();
            var oddTermGrouped = oddTermList.GroupBy(ip => ip.TypeOfWork)
                                            .Select(group => new { TypeOfWork = group.Key, TotalHours = group.Sum(ip => ip.Hours) })
                                            .ToList();


            foreach (var group in oddTermGrouped)
            {
                var existingRow = worksheet.RowsUsed().FirstOrDefault(s => s.Cell(8).Value.ToString() == group.TypeOfWork);
                if (existingRow == null)
                {
                    worksheet.Cell(row, 8).Value = group.TypeOfWork;
                    worksheet.Cell(row, col1).Value = group.TotalHours;
                    row++;
                }
                else
                {
                    existingRow.Cell(col1).Value = group.TotalHours;
                }
            }
            foreach (var group in evenTermGrouped)
            {
                var existingRow = worksheet.RowsUsed().FirstOrDefault(s => s.Cell(8).Value.ToString() == group.TypeOfWork);
                if (existingRow == null)
                {
                    worksheet.Cell(row, 8).Value = group.TypeOfWork;
                    worksheet.Cell(row, col2).Value = group.TotalHours;
                    row++;
                }
                else
                {
                    existingRow.Cell(col2).Value = group.TotalHours;
                }

            }
        }
        private void SumTables(IXLWorksheet worksheet, int row, int r, int col1, int col2, int col3)
        {
            double? sum;
            for (int i = r; i < row; i++)
            {
                var range1 = string.IsNullOrEmpty(worksheet.Row(i).Cell(col1).Value.ToString()) ? 0 : Convert.ToDouble(worksheet.Row(i).Cell(col1).Value.ToString());
                var range2 = string.IsNullOrEmpty(worksheet.Row(i).Cell(col2).Value.ToString()) ? 0 : Convert.ToDouble(worksheet.Row(i).Cell(col2).Value.ToString());
                sum = Convert.ToDouble(range1) + Convert.ToDouble(range2);
                worksheet.Cell(i, col3).Value = sum;
            }
            worksheet.Cell(row, 8).Value = "Итого";
            //MessageBox.Show(worksheet.Cell(3, col1).Value.ToString());
            //MessageBox.Show(worksheet.Cell(row, col1).Value.ToString());
            var range = worksheet.Range(3, col1, row, col1);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col1).Value = sum;
            worksheet.Cell(row, col1).Style.NumberFormat.Format = "0.##";
            range = worksheet.Range(3, col2, row, col2);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col2).Value = sum;
            worksheet.Cell(row, col2).Style.NumberFormat.Format = "0.##";
            range = worksheet.Range(3, col3, row, col3);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col3).Value = sum;
            worksheet.Cell(row, col3).Style.NumberFormat.Format = "0.##";
            for (int j = 8; j<15; j++)
            {
                worksheet.Cell(row, j).Style.Font.SetBold(true);
            }

        }

        private void SumAllTables(IXLWorksheet worksheet, int row, int r)
        {
            SumTables(worksheet, row, r, 9, 11, 13);
            SumTables(worksheet, row, r, 10, 12, 14);
        }
    }
}
