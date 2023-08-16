using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace Uniotech_LedTest.SubClass
{
    public class PFValues
    {
        private float min;
        private float max;
        private float error;

        public float Min { get => min; set => min = value; }
        public float Max { get => max; set => max = value; }
        public float Error { get => error; set => error = value; }
    }
}
