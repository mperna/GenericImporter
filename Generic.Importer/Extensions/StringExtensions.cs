using System;

namespace Generic.Importer.Extensions
{
    public static class StringExtensions
    {
        public static int ToInt32(this string value)
        {
            int retVal = 0;
            Int32.TryParse(value, out retVal);

            return retVal;
        }

        public static double ToDouble(this string value)
        {
            double retVal = 0;
            Double.TryParse(value, out retVal);

            return retVal;
        }

        public static decimal ToDecimal(this string value)
        {
            decimal retVal = 0;
            Decimal.TryParse(value, out retVal);

            return retVal;
        }
    }
}
