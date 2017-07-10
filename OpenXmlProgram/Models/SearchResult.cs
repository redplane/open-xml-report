namespace OpenXmlProgram.Models
{
    public class SearchResult<T>
    {
        public T Records { get; set; }

        public bool HasMoreRecords { get; set; }
    }
}