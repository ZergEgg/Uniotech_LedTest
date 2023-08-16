using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Uniotech_LedTest.SubClass
{
    internal class ZergEggParser
    {
        public struct StatusStruct
        {
            public byte MeasurementMode;
            public float Ch1Value;
            public float Ch2Value;
            public float Ch3Value;
            public float Ch4Value;
        }

        public struct ADCSettingStruct
        {
            public byte ADCCh;
            public UInt16 ADCSps;
            public UInt16 ADCNum;
            public UInt16 ADCWaitTime;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct BToStatus
        {
            [FieldOffset(0)]
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 255)]
            public byte[] bVal;

            [FieldOffset(0)]
            public StatusStruct Status;
        }

        [StructLayoutAttribute(LayoutKind.Explicit)]
        public struct ADCSetting
        {
            [FieldOffset(0)]
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 7)]
            public byte[] bVal;

            [FieldOffset(0)]
            public byte CH;
            [FieldOffset(1)]
            public UInt16 SPS;
            [FieldOffset(3)]
            public UInt16 SAMPLES;
            [FieldOffset(5)]
            public UInt16 NUMS;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct BToUint16
        {
            [FieldOffset(0)]
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
            public byte[] bVal;

            [FieldOffset(0)]
            public UInt16 Setting;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct BToFloat
        {
            [FieldOffset(0)]
            public byte b1;
            [FieldOffset(1)]
            public byte b2;
            [FieldOffset(2)]
            public byte b3;
            [FieldOffset(3)]
            public byte b4;

            [FieldOffset(0)]
            public float fVal;
        }
    }
}
