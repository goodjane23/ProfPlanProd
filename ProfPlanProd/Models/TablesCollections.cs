using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlanProd.Models
{
    internal static class TablesCollections
    {
        private static ObservableCollection<TableCollection> TablesCollection = new ObservableCollection<TableCollection>();

        public static ObservableCollection<TableCollection> GetTablesCollection()
        {
            return TablesCollection;    
        }

        public static void Clear()
        {
            TablesCollection.Clear();
        }

        public static void Add(TableCollection tabCol)
        {
            int foundIndex = GetTableIndexByName(tabCol.Tablename);
            if (foundIndex != -1)
            {
                TablesCollection[foundIndex] = tabCol;
            }
            else
            {
                TablesCollection.Add(tabCol);
            }
        }

        public static int GetTableIndexByName(string tableName, int selectedIndex = -2)
        {

            if (selectedIndex == -1)
            {
                return -1; 
            }

            for (int i = 0; i < TablesCollection.Count; i++)
            {
                if (TablesCollection[i].Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                {
                    return i; 
                }
            }

            return -1;
        }

        public static bool GetTableByName(string tableName, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                foreach (TableCollection table in TablesCollections.GetTablesCollectionWithP())
                {
                    if (table.Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            else if (selectedIndex == 1)
            {
                foreach (TableCollection table in TablesCollections.GetTablesCollectionWithF())
                {
                    if (table.Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            return false;

        }

        public static void SortTablesCollection()
        {
            var sortedCollectionP = TablesCollections.GetTablesCollectionWithP().Where(tc => tc.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1).OrderBy(tc =>
            {
                if (tc.Tablename.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) != -1) return 0;
                if (tc.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1) return 1;
                return 3;
            }).ThenBy(tc => tc.Tablename).ToList();

            var sortedCollectionF = TablesCollections.GetTablesCollectionWithF().OrderBy(tc =>
            {
                if (tc.Tablename.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) != -1) return 0;
                if (tc.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1) return 1;
                if (tc.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1) return 2;
                return 3;
            }).ThenBy(tc => tc.Tablename).ToList();

            TablesCollection.Clear();
            foreach (var table in sortedCollectionP)
            {
                TablesCollection.Add(table);
            }
            foreach (var table in sortedCollectionF)
            {
                TablesCollection.Add(table);
            }
        }

        public static ObservableCollection<TableCollection> GetTablesCollectionWithP()
        {
            return new ObservableCollection<TableCollection>(
                TablesCollection.Where(tc => tc.Tablename.StartsWith("П_")).ToList());
        }

        public static ObservableCollection<TableCollection> GetTablesCollectionWithF()
        {
            return new ObservableCollection<TableCollection>(
                TablesCollection.Where(tc => tc.Tablename.StartsWith("Ф_") || tc.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) != -1).ToList());
        }

        public static void AddInOldTabCol(TableCollection tabCol)
        {
            int foundIndex = GetTableIndexByName(tabCol.Tablename);

            if (foundIndex != -1)
            {
                for (int i = 0; i<tabCol.ExcelDataList.Count; i++)
                {
                    TablesCollection[foundIndex].ExcelDataList.Add(tabCol.ExcelDataList[i]);
                }
            }
            else
            {
                TablesCollection.Add(tabCol);
            }
        }

        public static void RemoveTableAtIndex(int index)
        {
            if (index >= 0 && index < TablesCollection.Count)
            {
                TablesCollection[index].ExcelDataList.Clear();
            }
        }

        public static void AddByIndex(int index, ExcelData tab)
        {
            TablesCollection[index].ExcelDataList.Add(tab);
        }

        public static int GetTableIndexForGenerate(string tableName, int selectedIndex)
        {

            if (selectedIndex == -1)
            {
                return -1; 
            }
            if (selectedIndex == 0)
            {
                for (int i = 0; i < TablesCollections.GetTablesCollectionWithP().Count; i++)
                {
                    if (TablesCollections.GetTablesCollectionWithP()[i].Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return i; 
                    }
                }
            }
            else
            {
                for (int i = 0; i < TablesCollections.GetTablesCollectionWithF().Count; i++)
                {
                    if (TablesCollections.GetTablesCollectionWithF()[i].Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return i; 
                    }
                }
            }

            return -1;

        }

    }
}
