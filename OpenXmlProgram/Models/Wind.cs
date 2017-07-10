using System;

namespace OpenXmlProgram.Models
{
    public class Wind
    {
        public string Id { get; set; }
        
        public string Time { get; set; } 

        public double Speed { get; set; }

        public double Direction { get; set; }

        public double Sm2m { get; set; }

        public double Dm2m { get; set; }

        public double Sm2s { get; set; }

        public double Dm2s { get; set; }

        public DateTime DateTime
        {
            get
            {
                try
                {
                    return DateTime.Parse(Time);
                }
                catch (Exception e)
                {
                    return DateTime.Now;
                }
            }
        }
    }
}