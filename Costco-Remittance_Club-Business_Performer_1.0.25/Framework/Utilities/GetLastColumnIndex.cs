using System;
using System.Collections.Generic;
using CostcoRemittance_ClubBusiness_Performer;
using UiPath.CodedWorkflows;

namespace TestFam
{
    public class GetLastColumnIndex : CodedWorkflow
    {
        [Workflow]
        public string Execute(int columnIndex)
        {
            string columnName = GetExcelColumnName(columnIndex);
            
            return columnName;
        }

        // Method to convert column index to Excel column name
        private string GetExcelColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                columnIndex--;  // Decrease by 1 to make it 0-indexed
                columnName = (char)(65 + (columnIndex % 26)) + columnName;  // 65 is ASCII code for 'A'
                columnIndex /= 26;  // Move to the next set of letters
            }
            return columnName;
        }
    }
}
