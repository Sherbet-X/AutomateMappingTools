using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomateMappingTool
{
    class ChangeFormat
    {
        public string formatDate(string date)
        {
            double d;
            DateTime dDate;

            if (DateTime.TryParse(date, out dDate))
            {
                date = dDate.ToString("dd/MM/yyyy");
            }
            else
            {
                if (double.TryParse(date, out d))
                {
                    dDate = DateTime.FromOADate(d);
                    date = dDate.ToString("dd/MM/yyyy");
                }
                else if (date == "-")
                {
                    date = null;
                }
                else
                {
                    date = "Invalid";
                }
            }

            return date;
        }

        public string formatSpeed(string speed)
        {
            if (speed.Contains("M"))
            {
                speed = speed.Replace("M", "");
                int convertSpeed = (Convert.ToInt32(speed)) * 1024;
                speed = convertSpeed.ToString();
            }
            else if (speed.Contains("G"))
            {
                speed = speed.Replace("G", "");
                int convertSpeed = Convert.ToInt32(speed) * 1024000;
                speed = convertSpeed.ToString();
            }
            else if (speed.Contains("K"))
            {
                speed = speed.Replace("K", "");
            }

            return speed;

        }
    }
}
