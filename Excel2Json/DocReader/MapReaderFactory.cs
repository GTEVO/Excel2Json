using System;

namespace Excel2Json.DocReader
{
    public static class MapReaderFactory
    {
        public static DocReader Create(string docType)
        {
            switch (docType) {
                case "{}":
                    return new ObjectReader();
                case "<>":
                    return new MapReader();
                case "[]":
                    return new ArrayReader();
                default:
                    throw new NotSupportedException($"{docType}");
            }
        }
    }
}