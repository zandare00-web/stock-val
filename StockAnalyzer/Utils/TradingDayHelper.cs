using System;
using System.Collections.Generic;

namespace StockAnalyzer.Utils
{
    public static class TradingDayHelper
    {
        private static readonly HashSet<DateTime> _holidays = new HashSet<DateTime>
        {
            // 2025
            new DateTime(2025,1,1), new DateTime(2025,1,28), new DateTime(2025,1,29), new DateTime(2025,1,30),
            new DateTime(2025,3,1), new DateTime(2025,5,5), new DateTime(2025,5,6), new DateTime(2025,6,6),
            new DateTime(2025,8,15), new DateTime(2025,10,3), new DateTime(2025,10,9),
            new DateTime(2025,12,25),
            // 2026
            new DateTime(2026,1,1),
            new DateTime(2026,2,16), new DateTime(2026,2,17), new DateTime(2026,2,18),  // 설날
            new DateTime(2026,3,1),   // 삼일절
            new DateTime(2026,5,5),   // 어린이날
            new DateTime(2026,5,24),  // 부처님오신날 (음력 4/8)
            new DateTime(2026,6,6),   // 현충일
            new DateTime(2026,8,15),  // 광복절
            new DateTime(2026,9,24), new DateTime(2026,9,25), new DateTime(2026,9,26),  // 추석
            new DateTime(2026,10,3),  // 개천절
            new DateTime(2026,10,9),  // 한글날
            new DateTime(2026,12,25), // 성탄절
        };

        public static bool IsTradingDay(DateTime d)
            => d.DayOfWeek != DayOfWeek.Saturday &&
               d.DayOfWeek != DayOfWeek.Sunday   &&
               !_holidays.Contains(d.Date);

        public static DateTime GetPreviousTradingDay(DateTime fromInclusive)
        {
            var d = fromInclusive.Date;
            while (!IsTradingDay(d)) d = d.AddDays(-1);
            return d;
        }

        public static DateTime GetLatestCompletedTradingDay(int marketCloseHourLocal = 18)
        {
            var d = DateTime.Today;
            if (DateTime.Now.Hour < marketCloseHourLocal) d = d.AddDays(-1);
            return GetPreviousTradingDay(d);
        }

        public static List<DateTime> GetRecentTradingDays(int count)
            => GetRecentTradingDays(count, DateTime.Today, includeIfTradingDay: true);

        public static List<DateTime> GetRecentTradingDays(int count, DateTime fromDate, bool includeIfTradingDay = true)
        {
            var list = new List<DateTime>();
            var d = includeIfTradingDay ? fromDate.Date : fromDate.Date.AddDays(-1);
            while (list.Count < count)
            {
                d = GetPreviousTradingDay(d);
                list.Add(d);
                d = d.AddDays(-1);
            }
            return list;  // [0]=최근, [count-1]=가장 오래된
        }

        public static string ToApiDate(DateTime d) => d.ToString("yyyyMMdd");
        public static string ToDisplayDate(DateTime d) => d.ToString("yyyy-MM-dd");
    }
}
