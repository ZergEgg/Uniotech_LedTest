using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Uniotech_LedTest.SubClass
{
    public class Sequence
    {
        public byte index { get; set; } = 0;
        public UInt16 outputValue { get; set; } = 0;
        public float result { get; set; } = 0;

        public Sequence(byte index, UInt16 outputValue, float result)
        {
            this.index = index;
            this.outputValue = outputValue;
            this.result = result;
        }
    }
}
