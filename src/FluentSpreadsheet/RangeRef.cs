namespace CodeCompanion.FluentSpreadsheet
{
    public struct RangeRef
    {
        public string Address { get; }

        public RangeRef(string address)
        {
            Address = address;
        }

        public static RangeRef From(IFluentRange range) => new(range.FullAddress);
    }
}