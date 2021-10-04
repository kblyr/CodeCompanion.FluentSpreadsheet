using System;
using System.Collections;
using System.Collections.Generic;

namespace CodeCompanion.FluentSpreadsheet
{
    public interface ISpreadsheetContext<TSharedData>
    {
        ISpreadsheetContext<TSharedData> WithSharedData(TSharedData data);
        ISpreadsheetContext<TSharedData> UpdateSharedData(Action<TSharedData> updateSharedData);
        ISpreadsheetContext<TSharedData> OnSheet(int sheetIndex);
        ISpreadsheetContext<TSharedData> OnSheet(string sheetName);
        ISpreadsheetContext<TSharedData> AddSheet(string sheetName, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> InsertSheet(string sheetName, int destinationIndex, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> CopySheet(int templateIndex, string name, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> CopySheet(int templateIndex, string name, int destinationIndex, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> CopySheet(string templateSheet, string name, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> CopySheet(string templateSheet, string name, int destinationIndex, bool setAsCurrent = true);
        ISpreadsheetContext<TSharedData> DeleteSheet(int sheetIndex);
        ISpreadsheetContext<TSharedData> DeleteSheet(string sheetName);
    }

    public abstract class SpreadsheetContextBase<TSharedData>
    {
        private readonly IWorkbookWrapper _workbook;

        private bool isSharedDataSet = false;
        private TSharedData sharedData = default;

        public abstract ISpreadsheetContext<TSharedData> Self { get; }
        protected abstract IWorksheetWrapper CreateWorksheet(IWorkbookWrapper workbook);

        public SpreadsheetContextBase(IWorkbookWrapper workbookWrapper)
        {
            _workbook = workbookWrapper;
        }

        public ISpreadsheetContext<TSharedData> WithSharedData(TSharedData data)
        {
            if (data is null)
                throw new ArgumentNullException("Shared data cannot be null");

            isSharedDataSet = true;
            sharedData = data;
            return Self;
        }

        public ISpreadsheetContext<TSharedData> UpdateSharedData(Action<TSharedData> updateSharedData)
        {
            updateSharedData.Invoke(sharedData);
            return Self;
        }

        public ISpreadsheetContext<TSharedData> OnSheet(int sheetIndex)
        {
            _workbook.SetCurrentWorksheet(sheetIndex);
            return Self;
        }

        public ISpreadsheetContext<TSharedData> OnSheet(string sheetName)
        {
            _workbook.SetCurrentWorksheet(sheetName);
            return Self;
        }

        public ISpreadsheetContext<TSharedData> AddSheet(string sheetName, bool setAsCurrent = true)
        {
            _workbook.AddWorksheet(sheetName);

            if (setAsCurrent)
                _workbook.SetCurrentWorksheet(sheetName);

            return Self;
        }

        public ISpreadsheetContext<TSharedData> InsertSheet(string sheetName, int destinationIndex, bool setAsCurrent = true)
        {
            var worksheet = CreateWorksheet(_workbook);
            _workbook.Worksheets.Insert(worksheet, destinationIndex);

            if (setAsCurrent)
                _workbook.SetCurrentWorksheet(sheetName);

            return Self;
        }
    }

    public interface IWorkbookWrapper
    {
        object Workbook { get; }
        IWorksheetWrapperEnumerable Worksheets { get; }
    }

    public interface IWorksheetWrapperEnumerable : IEnumerable<IWorksheetWrapper>
    {
        IWorkbookWrapper Workbook { get; }
        int Count { get; }
        void Add(IWorksheetWrapper worksheet);
        void Insert(IWorksheetWrapper worksheet, int destinationIndex);
        void Remove(IWorksheetWrapper worksheet);
    }

    public sealed class WorksheetWrapperEnumerable : IWorksheetWrapperEnumerable
    {
        private readonly List<IWorksheetWrapper> _worksheets = new();

        public IWorkbookWrapper Workbook { get; }

        public int Count => _worksheets.Count;

        public WorksheetWrapperEnumerable(IWorkbookWrapper workbook)
        {
            Workbook = workbook;
        }

        public void Add(IWorksheetWrapper worksheet)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<IWorksheetWrapper> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public void Insert(IWorksheetWrapper worksheet)
        {
            throw new NotImplementedException();
        }

        public void Remove(IWorksheetWrapper worksheet)
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    public interface IWorksheetWrapper
    {
        string Name { get; }
        object Worksheet { get; }
    }
}