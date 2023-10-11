using OfficeOpenXml;
using ReadExcelFiles.DTOs;

namespace ReadExcelFiles.Models
{
    public static class Excel
    {
        public static List<DateDto> ReturnExcelDatas()
        {
            string path = @"C:\Users\Admin\Downloads\2023.07.25-30_günlük 1.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(path));
            var workSheet = package.Workbook.Worksheets[0];
            var stations = new List<StationDto>();
            StationDto lastFoundStation = null;
            var dates = new List<DateDto>();
            DateDto lastDate = null;
            for (int row = 4; row <= workSheet.Dimension.End.Row; row++)
            {
                var counterOrStationNameOrDate = workSheet.Cells[row, 1].Value?.ToString();
                var weight = workSheet.Cells[row, 2].Value?.ToString();
                var pressureDifference = workSheet.Cells[row, 3].Value?.ToString();
                var pressure = workSheet.Cells[row, 4].Value?.ToString();
                var temprature = workSheet.Cells[row, 5].Value?.ToString();
                var spend = workSheet.Cells[row, 6].Value?.ToString();

                var counterOrStationNameTest = workSheet.Cells[row - 1, 1].Value?.ToString();
                var weightTest = workSheet.Cells[row - 1, 2].Value?.ToString();
                var pressureDifferenceTest = workSheet.Cells[row - 1, 3].Value?.ToString();
                var pressureTest = workSheet.Cells[row - 1, 4].Value?.ToString();
                var tempratureTest = workSheet.Cells[row - 1, 5].Value?.ToString();
                var spendTest = workSheet.Cells[row - 1, 6].Value?.ToString();
                if (DateTime.TryParse(counterOrStationNameOrDate, out DateTime date))
                {
                    lastDate = new DateDto() { Date = date };
                    dates.Add(lastDate);
                }
                else
                {
                    bool isFirstStation = counterOrStationNameOrDate != null && weight == null && pressure == null && pressureDifference == null && temprature == null && spend == null;
                    bool isSecondStation = counterOrStationNameTest != null && weightTest == null && pressureTest == null && pressureDifferenceTest == null && tempratureTest == null && spendTest == null;
                    if (isFirstStation && !isSecondStation)
                    {

                        lastFoundStation = new StationDto { Name = counterOrStationNameOrDate, SumAmount = 0 };
                        stations.Add(lastFoundStation);
                    }
                    if (isFirstStation && isSecondStation)
                    {
                        lastFoundStation.Name += counterOrStationNameOrDate;
                    }
                    if (counterOrStationNameOrDate != null && weight == null && pressure == null && pressureDifference == null && temprature != null && spend == null)
                    {
                        lastFoundStation.SumAmount = decimal.TryParse(temprature, out var sumAmountValue) ? sumAmountValue : null;
                    }
                    var gas = new GasDto()
                    {
                        CounterName = counterOrStationNameOrDate,
                        Weight = decimal.TryParse(weight, out var weightValue) ? weightValue : null,
                        PressureDifference = decimal.TryParse(pressureDifference, out var PressureDifferenceValue) ? PressureDifferenceValue : null,
                        Pressure = decimal.TryParse(pressure, out var pressureValue) ? pressureValue : null,
                        Temprature = decimal.TryParse(temprature, out var tempratureValue) ? tempratureValue : null,
                        Spend = decimal.TryParse(spend, out var spendValue) ? spendValue : null,
                    };
                    if (lastFoundStation != null && lastFoundStation.Counters == null)
                        lastFoundStation.Counters = new List<GasDto>();
                    if (lastFoundStation != null && weight != null && weight != null && pressure != null && pressureDifference != null && temprature != null && spend != null)
                        lastFoundStation.Counters.Add(gas);
                    if (lastDate != null)
                        lastDate.Stations = stations;
                }
            }
            return dates;
        }


        public static List<DateDto> ReturExcelDataByFilter(ExcelDataManipulationDto dto)
        {
            string path = @"C:\Users\Admin\Downloads\Book 6.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(path));
            var workSheet = package.Workbook.Worksheets[0];
            var stations = new List<StationDto>();
            StationDto lastFoundStation = null;
            var dates = new List<DateDto>();
            DateDto lastDate = null;

            for (int row = 4; row <= workSheet.Dimension.End.Row; row++)
            {
                var date = workSheet.Cells[row, 1].Value?.ToString();
                var sum = workSheet.Cells[row, 5].Value?.ToString();
                var counterName = workSheet.Cells[row, dto.counterNameId].Value?.ToString();
                var weight = workSheet.Cells[row, dto.weightId].Value?.ToString();
                var pressureDifference = workSheet.Cells[row, dto.pressureDifferenceId].Value?.ToString();
                var pressure = workSheet.Cells[row, dto.pressureId].Value?.ToString();
                var temprature = workSheet.Cells[row, dto.tempratureId].Value?.ToString();
                var spend = workSheet.Cells[row, dto.spendId].Value?.ToString();

                var previousDate = workSheet.Cells[row - 1, 1].Value?.ToString();
                var previousCounterName = workSheet.Cells[row - 1, dto.counterNameId].Value?.ToString();
                var previousWeight = workSheet.Cells[row - 1, dto.weightId].Value?.ToString();
                var previousPressureDifference = workSheet.Cells[row - 1, dto.pressureDifferenceId].Value?.ToString();
                var previousPressure = workSheet.Cells[row - 1, dto.pressureId].Value?.ToString();
                var previousTemprature= workSheet.Cells[row - 1, dto.tempratureId].Value?.ToString();
                var previuosSpend = workSheet.Cells[row - 1, dto.spendId].Value?.ToString();

                bool isFirstStation = !DateTime.TryParse(date, out DateTime modifingDate1) && date != null && counterName == null && pressure == null && pressureDifference == null && temprature == null && spend == null;
                bool isSecondStation = !DateTime.TryParse(previousDate, out DateTime modifingDate2) && previousDate != null && previousCounterName == null && previousPressure == null && previousPressureDifference == null && previousTemprature == null && previuosSpend == null;

                if (DateTime.TryParse(date, out DateTime modifingDate))
                {
                    lastDate = new DateDto() { Date = modifingDate };
                    dates.Add(lastDate);
                }
                else
                {
                    if (isFirstStation && !isSecondStation)
                    {

                        lastFoundStation = new StationDto { Name = date, SumAmount = 0 };
                        stations.Add(lastFoundStation);
                    }
                    if (isFirstStation && isSecondStation)
                    {
                        lastFoundStation.Name += date;
                    }
                    if (date != null && sum != null && (pressure == null || pressureDifference == null || spend == null || temprature == null))
                    {
                        lastFoundStation.SumAmount = decimal.TryParse(temprature, out var sumAmountValue) ? sumAmountValue : null;
                    }
                    var gas = new GasDto()
                    {
                        CounterName = counterName,
                        Weight = decimal.TryParse(weight, out var weightValue) ? weightValue : null,
                        PressureDifference = decimal.TryParse(pressureDifference, out var PressureDifferenceValue) ? PressureDifferenceValue : null,
                        Pressure = decimal.TryParse(pressure, out var pressureValue) ? pressureValue : null,
                        Temprature = decimal.TryParse(temprature, out var tempratureValue) ? tempratureValue : null,
                        Spend = decimal.TryParse(spend, out var spendValue) ? spendValue : null,
                    };
                    if (lastFoundStation != null && lastFoundStation.Counters == null)
                        lastFoundStation.Counters = new List<GasDto>();
                    if (lastFoundStation != null && weight != null && weight != null && pressure != null && pressureDifference != null && temprature != null && spend != null)
                        lastFoundStation.Counters.Add(gas);
                    if (lastDate != null)
                        lastDate.Stations = stations;
                }
            }
            return dates;
        }
    }
}
