using System;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelDataReader;

namespace Excel2Json
{
    public class ExcelWorksheet : IDisposable
    {
        static ExcelWorksheet()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public int EndColumn { get; private set; }
        public int EndRow { get; private set; }

        private readonly DataSet _dataSet;
        private readonly DataTable _sheet;

        private readonly Type _stringType = typeof(string);

        public ExcelWorksheet(FileInfo fileInfo)
        {
            using (var fileStream = fileInfo.Open(FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                using (var reader = ExcelReaderFactory.CreateReader(fileStream)) {
                    _dataSet = reader.AsDataSet();
                    _sheet = _dataSet.Tables[0];

                    EndColumn = _sheet.Columns.Count;
                    EndRow = _sheet.Rows.Count;
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool HasValue(int row, int col)
        {
            var value = _sheet.Rows[row - 1][col - 1];
            if (value == DBNull.Value || value is null) {
                return false;
            }
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public string GetString(int row, int col)
        {
            var value = _sheet.Rows[row - 1][col - 1];
            if (value == DBNull.Value || value is null) {
                return null;
            }
            return (string)Convert.ChangeType(value, _stringType);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public T GetValue<T>(int row, int col)
        {
            var value = _sheet.Rows[row - 1][col - 1];
            if (value == DBNull.Value || value is null) {
                return default;
            }
            return (T)Convert.ChangeType(value, typeof(T));
        }

        public void Dispose()
        {
            _dataSet.Dispose();
        }
    }
}
