using System;
using ExcelOrderAddIn.DateTime;

namespace Tests.Stubs
{
    public class DateTimeForTests : IDateTime
    {
        public DateTime Now { get; set; }
    }
}
