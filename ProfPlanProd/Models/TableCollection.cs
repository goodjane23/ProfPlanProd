using ProfPlanProd.ViewModels.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlanProd.Models
{
    internal class TableCollection : ViewModel, IEnumerable
    {
        private string _tablename = null;
        private ObservableCollection<ExcelData> _excelDataList = new ObservableCollection<ExcelData>();

        public ObservableCollection<ExcelData> ExcelDataList
        {
            get { return _excelDataList; }
            set
            {
                if (_excelDataList != value)
                {
                    _excelDataList = value;
                    OnPropertyChanged(nameof(ExcelDataList));
                }
            }
        }
        public string Tablename
        {
            get { return _tablename; }
            set
            {
                if (_tablename != value)
                {
                    _tablename = value;
                    OnPropertyChanged(nameof(Tablename));
                }
            }
        }

        public TableCollection(string tablename, ObservableCollection<ExcelData> col)
        {
            Tablename = tablename;
            ExcelDataList = col;
        }
        public TableCollection(string tablename)
        {
            Tablename = tablename;
            ExcelDataList = new ObservableCollection<ExcelData>();
        }
        public TableCollection()
        {
            ExcelDataList = new ObservableCollection<ExcelData>();
        }



        public void SubscribeToExcelDataChanges()
        {
            foreach (var excelModel in _excelDataList)
            {
                excelModel.PropertyChanged -= ExcelModel_PropertyChanged;
            }

            foreach (var excelModel in _excelDataList)
            {
                excelModel.PropertyChanged += ExcelModel_PropertyChanged;
            }

            UpdateHours();
        }
        public void ExcelModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateHours();
        }
        private void UpdateHours()
        {
            TotalHours = _excelDataList.OfType<ExcelModel>().Where(x => x.Total != null).Sum(x => Convert.ToDouble(x.Total));
            AutumnHours = _excelDataList.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals("нечет", StringComparison.OrdinalIgnoreCase))
                                .Sum(x => Convert.ToDouble(x.Total));
            SpringHours = _excelDataList.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals("чет", StringComparison.OrdinalIgnoreCase))
                                .Sum(x => Convert.ToDouble(x.Total));
        }
        private double _totalHours;
        public double TotalHours
        {
            get { return _totalHours; }
            set
            {
                if (_totalHours != value)
                {
                    _totalHours = value;
                    OnPropertyChanged(nameof(TotalHours));
                }
            }
        }
        private double _autumnHours;
        public double AutumnHours
        {
            get { return _autumnHours; }
            set
            {
                if (_autumnHours != value)
                {
                    _autumnHours = value;
                    OnPropertyChanged(nameof(AutumnHours));
                }
            }
        }
        private double _springHours;

        public double SpringHours
        {
            get { return _springHours; }
            set
            {
                if (_springHours != value)
                {
                    _springHours = value;
                    OnPropertyChanged(nameof(SpringHours));
                }
            }
        }

        public IEnumerator<ExcelData> GetEnumerator()
        {
            return _excelDataList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void SortExcelDataListByNumber()
        {
            List<ExcelModel> sortedList = new List<ExcelModel>();
            foreach(ExcelModel ex in  _excelDataList)
            {
                sortedList.Add(ex);
            }
            sortedList.Sort((x, y) => x.Number.CompareTo(y.Number));
            _excelDataList.Clear();
            foreach (var item in sortedList)
            {
                _excelDataList.Add(item);
            }
        }

    }
}
