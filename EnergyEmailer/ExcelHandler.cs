using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace EnergyEmailer
{
    [Serializable()]
    public class InvalidExcelEntryException : System.Exception
    {
        public string DataType
        {
            get;
            private set;
        }

        public string DataValue
        {
            get;
            private set;
        }

        public InvalidExcelEntryException(string dataType, string dataValue)
            : base("Could not extract data from Excel cell.")
        {
            DataType = dataType;
            DataValue = dataValue;
        }

        protected InvalidExcelEntryException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context) { }
    }

    public static class ExcelHandler
    {
        public static List<ExcelRow> GetWorksheet(Excel.Worksheet worksheet)
        {
            Excel.Range range = worksheet.UsedRange;
            int numEntries = range.Rows.Count;
            List<ExcelRow> rows = new List<ExcelRow>(numEntries);
            Object[,] valueArray = (Object[,])range.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            if (numEntries < 2)
            {
                MessageBox.Show("Worksheet is empty.", "Error");
                return null;
            }

            for (int i = 2; i < numEntries + 1; ++i)
            {
                try
                {
                    ExcelRow newRow = new ExcelRow(valueArray, i);
                    rows.Add(newRow);
                }
                catch (InvalidExcelEntryException ex)
                {
                    DialogResult skipRow = MessageBox.Show(
                        String.Format("An invalid data error occured while parsing.\n\nRow: {0}\nData type: {1}\nData value: {2}\n\nPress \"OK\" to skip this row and continue or \"Cancel\" to discontinue loading operation.", i.ToString(), ex.DataType, ex.DataValue),
                        "Error",
                        MessageBoxButtons.OKCancel);

                    if (skipRow == DialogResult.Cancel)
                        return null;
                }
                catch (Exception ex)
                {
                    DialogResult skipRow = MessageBox.Show(
                        String.Format("An unidentified error occured while parsing row {0}:\n\n{1}\n\nPress \"OK\" to skip this row and continue or \"Cancel\" to discontinue loading operation.", i.ToString(), ex.Message),
                        "Error",
                        MessageBoxButtons.OKCancel);

                    if (skipRow == DialogResult.Cancel)
                        return null;
                }
            }

            return rows;
        }
    }

    [Serializable()]
    public class ExcelRow
    {
        private const int COL_EMAIL_ADDRESS = 1;
        private const int COL_USER_GROUP = 2;
        private const int COL_ENERGY_YOU = 3;
        private const int COL_ENERGY_LOW = 4;
        private const int COL_ENERGY_HIGH = 5;
        private const int COL_RATING = 6;

        public string EmailAddress
        {
            get;
            private set;
        }

        public int UserGroupId
        {
            get;
            private set;
        }

        public double YourEnergyUse
        {
            get;
            private set;
        }

        public double HighestEnergyUse
        {
            get;
            private set;
        }

        public double LowestEnergyUse
        {
            get;
            private set;
        }

        public int Rating
        {
            get;
            private set;
        }

        public ExcelRow(Object[,] table, int rowNum)
        {
            try
            {
                EmailAddress = Convert.ToString(table[rowNum, COL_EMAIL_ADDRESS]);
                var addr = new System.Net.Mail.MailAddress(EmailAddress);
            }
            catch
            {
                throw new InvalidExcelEntryException("Email address", table[rowNum, COL_EMAIL_ADDRESS].ToString());
            }

            try
            {
                UserGroupId = Convert.ToInt32(table[rowNum, COL_USER_GROUP]);
                if (UserGroupId < 1 || UserGroupId > 3)
                {
                    throw new Exception();
                }
            }
            catch
            {
                throw new InvalidExcelEntryException("User group", table[rowNum, COL_USER_GROUP].ToString());
            }

            try
            {
                YourEnergyUse = Convert.ToDouble(table[rowNum, COL_ENERGY_YOU]);
            }
            catch
            {
                throw new InvalidExcelEntryException("Resident's energy use", table[rowNum, COL_ENERGY_YOU].ToString());
            }

            try
            {
                LowestEnergyUse = Convert.ToDouble(table[rowNum, COL_ENERGY_LOW]);
            }
            catch
            {
                throw new InvalidExcelEntryException("Compared energy use", table[rowNum, COL_ENERGY_LOW].ToString());
            }

            try
            {
                HighestEnergyUse = Convert.ToDouble(table[rowNum, COL_ENERGY_HIGH]);
            }
            catch
            {
                throw new InvalidExcelEntryException("Best energy use", table[rowNum, COL_ENERGY_HIGH].ToString());
            }

            try
            {
                Rating = Convert.ToInt32(table[rowNum, COL_RATING]);
                if (Rating < 1 || Rating > 1)
                {
                    throw new Exception();
                }
            }
            catch
            {
                throw new InvalidExcelEntryException("Rating", table[rowNum, COL_RATING].ToString());
            }
        }
    }
}
