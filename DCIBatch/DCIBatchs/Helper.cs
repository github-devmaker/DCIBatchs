using System;
using System.Globalization;

namespace DCIBatchs
{
    public class Helper
    {
        public string SetDigit(string value, int digit)
        {
            try
            {
                return Convert.ToInt32(value).ToString($"D{digit}");
            }
            catch
            {
                return value;
            }
        }

        public double ConvDecToDb(decimal val)
        {
            try
            {
                return Convert.ToDouble(val);
            }
            catch
            {
                return 0;
            }
        }
        public decimal ConvStrToDec(string val)
        {
            try
            {
                return val != "" ? Convert.ToDecimal(val) : 0;
            }
            catch
            {
                return 0;
            }
        }
        public double ConvStrToDB(string val)
        {
            try
            {
                return val != "" ? Convert.ToDouble(val) : 0;
            }
            catch
            {
                return 0;
            }
        }
        public int ConvStrToInt(string val)
        {
            try
            {
                return val != "" ? Convert.ToInt32(val) : 0;
            }
            catch
            {
                return 0;
            }
        }
        public string ConvInt2Str(int val)
        {
            try
            {
                return Convert.ToString(val);
            }
            catch
            {
                return "";
            }
        }
        public int ConvDecToInt(decimal? val)
        {
            try
            {
                return val != null ? (int)val : 0;
            }
            catch
            {
                return 0;
            }
        }
    }
}
