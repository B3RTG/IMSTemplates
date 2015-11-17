using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IMSClasses.Jobs
{
    public class Schelude
    {
        public enum periocity_values
        {
            Daily = 1,
            Monthly = 2,
            Weekly =3
        }
        public const char _SPLIT_CHAR_ = ',';

        public String Periodicity;
        public int Day;
        public String WeekDays;
        public int Minut;
        public int Hour;

        private List<DayOfWeek> Days;

        public Schelude()
        {
            this.Days = new List<DayOfWeek>();
            
        }


        public void fill_private_vars () {
            if (this.WeekDays != null && !this.WeekDays.Equals(String.Empty))
            {
                foreach (String sDay in this.WeekDays.Split(_SPLIT_CHAR_))
                {
                    switch (sDay.ToUpper()) {
                        case "LUNES":
                            this.Days.Add(DayOfWeek.Monday);
                            break;
                        case "MARTES":
                            this.Days.Add(DayOfWeek.Tuesday);
                            break;
                        case "MIERCOLES":
                            this.Days.Add(DayOfWeek.Wednesday);
                            break;
                        case "JUEVES":
                            this.Days.Add(DayOfWeek.Thursday);
                            break;
                        case "VIERNES":
                            this.Days.Add(DayOfWeek.Friday);
                            break;
                        case "SABADO":
                            this.Days.Add(DayOfWeek.Saturday);
                            break;
                        case "DOMINGO":
                            this.Days.Add(DayOfWeek.Sunday);
                            break;
                    }

                }
            }
        }

    }
}
