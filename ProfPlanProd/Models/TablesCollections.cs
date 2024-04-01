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
    }
}
