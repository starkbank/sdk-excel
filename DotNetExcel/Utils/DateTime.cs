using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StarkBankMVP
{
    internal class StarkDate
    {
        public DateTime? Value;

        public StarkDate(DateTime? value)
        {
            Value = value;
        }

        public StarkDate(string value)
        {
            Value = DateTime.Parse(value);
        }

        public override string ToString()
        {
            if (Value == null)
                return null;
            DateTime value = (DateTime)Value;
            return value.ToString("yyyy-MM-dd");
        }
    }

    internal class StarkDateTime
    {
        public DateTime? Value;

        public StarkDateTime(DateTime? value)
        {
            Value = value;
        }
        public StarkDateTime(string value)
        {
            Value = DateTime.Parse(value);
        }

        public override string ToString()
        {
            if (Value == null)
                return null;
            DateTime value = (DateTime)Value;
            return value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffffff") + "+00:00";
        }
    }
}
