using System;
using System.Collections.Generic;
using System.Data;
using CostcoRemittance_ClubBusiness_Performer;
using UiPath.CodedWorkflows;

namespace TestFam
{
    public class GetActionCenterBatches : CodedWorkflow
    {
        [Workflow]
        public List<DataTable> Execute(DataTable fullData, int batchRows = 500)
        {
            List<DataTable> batches = new List<DataTable>();

            // Edge Case 1 & 2: Null or Empty Table
            if (fullData == null || fullData.Rows.Count == 0)
            {
                return batches;
            }

            // Edge Case 6: Invalid batch size
            if (batchRows <= 0)
            {
                throw new ArgumentException("Batch size must be greater than 0.");
            }

            int totalRows = fullData.Rows.Count;
            int currentIndex = 0;

            while (currentIndex < totalRows)
            {
                DataTable batchTable = new DataTable();

                // Add Index column as first
                batchTable.Columns.Add("Index", typeof(int));

                // Add the rest of the columns from fullData
                foreach (DataColumn col in fullData.Columns)
                {
                    // Avoid duplicate column names
                    string columnName = col.ColumnName == "Index" ? "Original_Index" : col.ColumnName;
                    batchTable.Columns.Add(columnName, col.DataType);
                }

                int rowsToCopy = Math.Min(batchRows, totalRows - currentIndex);

                for (int i = 0; i < rowsToCopy; i++)
                {
                    DataRow originalRow = fullData.Rows[currentIndex + i];
                    DataRow newRow = batchTable.NewRow();

                    newRow["Index"] = i + 1;

                    foreach (DataColumn col in fullData.Columns)
                    {
                        string colName = col.ColumnName == "Index" ? "Original_Index" : col.ColumnName;
                        newRow[colName] = originalRow[col.ColumnName];
                    }

                    batchTable.Rows.Add(newRow);
                }

                batches.Add(batchTable);
                currentIndex += rowsToCopy;
            }

            return batches;
        }
    }
}
