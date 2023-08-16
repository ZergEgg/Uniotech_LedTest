using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Uniotech_LedTest.SubClass
{
    public delegate void ZergEggDataRecvEventHandler(byte[] arr, int len);

    internal class ZergEggSerialPort
    {
        public ZergEggDataRecvEventHandler zergEggDataRecvEventHandler;

        SerialPort sp;

        String port;
        int baudRate;

        byte[] oneLineArr = new byte[32767];
        int recvCnt = 0;

        public ZergEggSerialPort(String port, int baudRate)
        {
            this.port = port;
            this.baudRate = baudRate;
        }

        public bool Connect()
        {
            try
            {
                sp = new SerialPort(this.port, this.baudRate);
                sp.DataReceived += SerialPortDataRecv;
                sp.Open();

                return true;
            }
            catch(Exception e) 
            {
                return false;
            }
            
        }

        public void Disconnect()
        {
            sp.DiscardInBuffer();
            sp.DataReceived -= SerialPortDataRecv;
            sp.Close();
        }

        public bool IsConnected()
        {
            return sp.IsOpen;
        }

        private void SerialPortDataRecv(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                SerialReceiver(sender, e);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void SerialReceiver(object sender, EventArgs e)
        {
            int cnt = sp.BytesToRead;

            for (int i = 0; i < cnt; i++)
            {
                byte b = (byte)sp.ReadByte();
                switch (b)
                {
                    case 0xC0:
                        recvCnt = 0;
                        oneLineArr[recvCnt++] = b;
                        break;
                    case 0xC2:
                        oneLineArr[recvCnt++] = b;
                        MakeOneLine();
                        recvCnt = 0;
                        break;
                    default:
                        oneLineArr[recvCnt++] = b;
                        break;
                }
            }
        }

        private void MakeOneLine()
        {
            byte[] arr = new byte[recvCnt];

            int i = 0;
            int offsetCnt = 0;

            arr[i++] = 0xC0;

            for (; i < recvCnt - 1; i++)
            {
                byte b = oneLineArr[i + offsetCnt];
                switch (b)
                {
                    case 0xDB:
                        offsetCnt++;
                        switch (oneLineArr[i + offsetCnt])
                        {
                            case 0xDC:
                                arr[i] = 0xC0;
                                break;
                            case 0xDD:
                                arr[i] = 0xDB;
                                break;
                            case 0xDE:
                                arr[i] = 0xC2;
                                break;
                        }
                        break;
                    default:
                        arr[i] = b;
                        break;
                }
            }

            if (CalcChkSum(arr, i - offsetCnt) == true)
            {
                zergEggDataRecvEventHandler(arr, i);
            }
        }

        private bool CalcChkSum(byte[] arr, int len)
        {
            bool rtnVal = true;
            byte ChkSum = 0xC2;

            for (int i = 0; i < len - 1; i++)
            {
                ChkSum += arr[i];
            }

            if (ChkSum != arr[len - 1])
            {
                return false;
            }

            return rtnVal;
        }

        public void SendData(byte cmd1, byte cmd2, byte cmd3, byte[] data, int len)
        {
            byte[] arr = new byte[len + 8];
            byte chkSum = 0xC2;

            arr[0] = 0xC0;
            arr[1] = (byte)(((len) >> 8) & 0xFF);
            arr[2] = (byte)(((len)) & 0xFF);
            arr[3] = cmd1;
            arr[4] = cmd2;
            arr[5] = cmd3;

            for (int i = 0; i < len; i++)
            {
                arr[6 + i] = data[i];
            }

            for (int i = 0; i < len + 6; i++)
            {
                chkSum += arr[i];
            }

            arr[6 + len] = chkSum;
            arr[6 + len + 1] = 0xC2;

            SendArr(arr, len + 8);
        }

        public void SendArr(byte[] arr, int len)
        {
            byte[] sendArr = new byte[len + 100];
            int spCnt = 0;

            sendArr[0] = 0xC0;

            for (int i = 1; i < len; i++)
            {
                switch (arr[i])
                {
                    case 0xC0:
                        sendArr[i + spCnt] = 0xDB;
                        spCnt++;
                        sendArr[i + spCnt] = 0xDC;
                        break;
                    case 0xDB:
                        sendArr[i + spCnt] = 0xDB;
                        spCnt++;
                        sendArr[i + spCnt] = 0xDD;
                        break;
                    case 0xC2:
                        sendArr[i + spCnt] = 0xDB;
                        spCnt++;
                        sendArr[i + spCnt] = 0xDE;
                        break;
                    default:
                        sendArr[i + spCnt] = arr[i];
                        break;
                }
            }

            sendArr[len + spCnt] = 0xC2;

            sp.Write(sendArr, 0, len + spCnt + 1);
        }
    }
}
