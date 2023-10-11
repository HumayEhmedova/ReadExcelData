namespace ReadExcelFiles.DTOs
{
    public class StationDto
    {
        public string? Name { get; set; }
        public decimal? SumAmount { get; set; }
        public List<GasDto> Counters { get; set; }
    }
}
 