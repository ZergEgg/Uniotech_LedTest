using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using Uniotech_LedTest.SubClass;
using static Uniotech_LedTest.SubClass.ZergEggParser;
using System.IO;
using System.Threading;

namespace Uniotech_LedTest
{
    public partial class MainForm : Form
    {
        TextBox targetTextBox;

        ZergEggSerialPort sp;

        List<Sequence> seqList = new List<Sequence>();
        int opMode = 1;

        static string folderPath = System.Windows.Forms.Application.StartupPath;
        static string mode2FileName = Path.Combine(folderPath, "Mode2_Result.csv");
        static String mode3FileName = Path.Combine(folderPath, "Mode3_Result.csv");


        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadSetting();

            GetPortName();

            InitSequence();

            DisplayInit();

            //CreateExcelInstance();
        }

        private StreamWriter CreateExcelInstance(String filePath)
        {
            try
            {
                StreamWriter sw = null;
                if (File.Exists(filePath))
                {
                    sw = File.AppendText(filePath);
                }
                else
                {
                    sw = File.CreateText(filePath);

                    if (filePath == mode2FileName)
                    {
                        sw = MakeHeader(sw, 2);
                    }
                    else if (filePath == mode3FileName)
                    {
                        sw = MakeHeader(sw, 3);
                    }
                }

                return sw;

                //WriteToExcel("aaaa", new float[] { 1.2f, 1.3f, 1.4f, 11.5f });
            }
            catch (Exception ex)
            {
                CloseExcel();
            }

            return null;
        }

        private StreamWriter MakeHeader(StreamWriter sw, int mode)
        {
            switch (mode)
            {
                case 2:
                    sw.WriteLine("TIME,S/N,BLUE,GREEN,RED,ORANGE,PASS/FAIL");
                    break;
                case 3:
                    sw.WriteLine("TIME,S/N,L510.MIN,L510.MAX,L580.MIN,L580.MAX,L670.MIN,L670.MAX,L725.MIN,L725.MAX");
                    break;
            }

            return sw;
        }

        private void WriteToExcelMode2(String sn, float[] values, float[] min, float[] max, float[] err)
        {
            StreamWriter sw = CreateExcelInstance(mode2FileName);

            if(textBoxMode2Ch0Result.Text.Length == 0)
            {
                MessageBox.Show("측정되지 않았습니다. 확인해주세요.");
                return;
            }

            String result = "Pass";

            for(int i = 0; i < 4; i++)
            {
                if (values[i] < min[i]  || values[i] > max[i])
                {
                    result = "Fail";
                    break;
                }
            }

            String appendText =
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + sn + "," +
                textBoxMode2Ch0Result.Text + "," + textBoxMode2Ch1Result.Text + "," +
                textBoxMode2Ch2Result.Text + "," + textBoxMode2Ch3Result.Text + "," +
                result;

            sw.WriteLine(appendText);

            sw.Flush();
            sw.Close();
        }

        private void WriteToExcelMode3()
        {
            if (textBoxSn.Text.Length == 0)
            {
                MessageBox.Show("시리얼 번호가 입력되어 있지 않습니다.");
                return;
            }

            if (textBoxMode3Ch1ResultMin.Text.Length == 0 || textBoxMode3Ch1ResultMax.Text.Length == 0 || textBoxMode3Ch2ResultMin.Text.Length == 0 || textBoxMode3Ch2ResultMax.Text.Length == 0 ||
               textBoxMode3Ch3ResultMin.Text.Length == 0 || textBoxMode3Ch3ResultMax.Text.Length == 0 || textBoxMode3Ch4ResultMin.Text.Length == 0 || textBoxMode3Ch4ResultMax.Text.Length == 0)
            {
                MessageBox.Show("모든값이 측정되지 않았습니다. 확인해주세요.");
                return;
            }

            String appendText =
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + textBoxSn.Text + "," +
                textBoxMode3Ch1ResultMin.Text + "," + textBoxMode3Ch1ResultMax.Text + "," +
                textBoxMode3Ch2ResultMin.Text + "," + textBoxMode3Ch2ResultMax.Text + "," +
                textBoxMode3Ch3ResultMin.Text + "," + textBoxMode3Ch3ResultMax.Text + "," +
                textBoxMode3Ch4ResultMin.Text + "," + textBoxMode3Ch4ResultMax.Text;

            StreamWriter sw = CreateExcelInstance(mode3FileName);
            sw.WriteLine(appendText);

            sw.Flush();
            sw.Close();
        }

        private bool CheckValue(float val, float min, float max)
        {
            if (val >= min && val <= max)
            {
                return true;
            }

            return false;
        }

        private void SetDefault()
        {
            Properties.Settings.Default.CH1_VALUE = 1080;
            Properties.Settings.Default.CH2_VALUE = 300;
            Properties.Settings.Default.CH3_VALUE = 1497;
            Properties.Settings.Default.CH4_VALUE = 600;

            Properties.Settings.Default.CH1_PF_MAX = -121;
            Properties.Settings.Default.CH1_PF_MIN = -189;
            Properties.Settings.Default.CH2_PF_MAX = -28.8f;
            Properties.Settings.Default.CH2_PF_MIN = -33;
            Properties.Settings.Default.CH3_PF_MAX = -174;
            Properties.Settings.Default.CH3_PF_MIN = -241.5f;
            Properties.Settings.Default.CH4_PF_MAX = -61;
            Properties.Settings.Default.CH4_PF_MIN = -90.3f;

            Properties.Settings.Default.CH1_PF_ERROR = 0.5f;
            Properties.Settings.Default.CH2_PF_ERROR = 0.5f;
            Properties.Settings.Default.CH3_PF_ERROR = 0.5f;
            Properties.Settings.Default.CH4_PF_ERROR = 0.5f;

            Properties.Settings.Default.Save();
        }

        private void InitSequence()
        {
            seqList = new List<Sequence>();
        }

        private void LoadSetting()
        {
        }

        private void SaveSetting()
        {
        }

        private void GetPortName()
        {
            comboBoxPorts.Items.Clear();

            foreach (String port in SerialPort.GetPortNames())
            {
                comboBoxPorts.Items.Add(port);

                comboBoxPorts.SelectedIndex = 0;
            }
        }

        private void DisplayInit()
        {
            numericUpDownCh0Setting.Value = Properties.Settings.Default.CH1_VALUE;
            numericUpDownCh1Setting.Value = Properties.Settings.Default.CH2_VALUE;
            numericUpDownCh2Setting.Value = Properties.Settings.Default.CH3_VALUE;
            numericUpDownCh3Setting.Value = Properties.Settings.Default.CH4_VALUE;

            numericUpDownPFCH0Min.Value = (decimal)Properties.Settings.Default.CH1_PF_MIN;
            numericUpDownPFCH1Min.Value = (decimal)Properties.Settings.Default.CH2_PF_MIN;
            numericUpDownPFCH2Min.Value = (decimal)Properties.Settings.Default.CH3_PF_MIN;
            numericUpDownPFCH3Min.Value = (decimal)Properties.Settings.Default.CH4_PF_MIN;

            numericUpDownPFCH0Max.Value = (decimal)Properties.Settings.Default.CH1_PF_MAX;
            numericUpDownPFCH1Max.Value = (decimal)Properties.Settings.Default.CH2_PF_MAX;
            numericUpDownPFCH2Max.Value = (decimal)Properties.Settings.Default.CH3_PF_MAX;
            numericUpDownPFCH3Max.Value = (decimal)Properties.Settings.Default.CH4_PF_MAX;

            numericUpDownPFCH0Error.Value = (decimal)Properties.Settings.Default.CH1_PF_ERROR;
            numericUpDownPFCH1Error.Value = (decimal)Properties.Settings.Default.CH2_PF_ERROR;
            numericUpDownPFCH2Error.Value = (decimal)Properties.Settings.Default.CH3_PF_ERROR;
            numericUpDownPFCH3Error.Value = (decimal)Properties.Settings.Default.CH4_PF_ERROR;

            comboBoxADCCh.SelectedIndex = 0;
            comboBoxSPS.SelectedIndex = 0;
        }

        private void buttonConn_Click(object sender, EventArgs e)
        {
            if (sp == null || sp.IsConnected() == false)
            {
                if(comboBoxPorts.SelectedIndex < 0)
                {
                    MessageBox.Show("포트를 선택해주세요.");
                    return;
                }

                sp = new ZergEggSerialPort(comboBoxPorts.SelectedItem.ToString(), 115200);
                sp.zergEggDataRecvEventHandler += DataRecvEventHandler;

                if (sp.Connect())
                {
                    SendWhoAmI();
                    //GetADCSetting();
                    //GetDACSetting();
                    buttonConn.Text = "끊기";
                }
            }
            else
            {
                sp.Disconnect();
                buttonConn.Text = "연결";
            }
        }

        private void GetADCSetting()
        {
            SendData(0x01, 0x02, 0x02, new byte[1] { 0x00 }, 1);
        }

        private void GetDACSetting()
        {
            SendData(0x01, 0x01, 0x02, new byte[1] { 0x00 }, 1);
        }

        private void SendWhoAmI()
        {
            SendData(0x55, 0xAA, 0x00, new byte[1] { 0x00 }, 1);
        }

        private void DataRecvEventHandler(byte[] arr, int len)
        {
            byte cmd1 = arr[3];

            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate ()
                {
                    switch (cmd1)
                    {
                        case 0x55:
                            WhoAmIParser(arr, len);
                            break;
                        case 0x01: // Setting
                            Cmd1Parser(arr, len);
                            break;
                        case 0x02: // Measurement
                            Cmd2Parser(arr, len);
                            break;
                        case 0x03:
                            Cmd3Parser(arr, len);
                            break;
                    }
                }));
            }
            else
            {
                switch (cmd1)
                {
                    case 0x55:
                        WhoAmIParser(arr, len);
                        break;
                    case 0x01: // Setting
                        Cmd1Parser(arr, len);
                        break;
                    case 0x02: // Measurement
                        Cmd2Parser(arr, len);
                        break;
                    case 0x03:
                        Cmd3Parser(arr, len);
                        break;
                }
            }


        }

        private void WhoAmIParser(byte[] arr, int len)
        {
            byte cmd2 = arr[4];
            byte cmd3 = arr[5];
            byte[] data = new byte[len];
            Array.Copy(arr, 6, data, 0, len - 6);

            byte returnData = arr[6];
            byte ch = arr[7];
            UInt16 sps = (UInt16)((arr[8] << 8) | arr[9]);
            UInt16 samples = (UInt16)((arr[10] << 8) | arr[11]);

            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate ()
                {
                    textBoxADCCh.Text = comboBoxADCCh.Items[ch].ToString();
                    textBoxADCSps.Text = comboBoxSPS.Items[sps].ToString();
                    textBoxADCSamples.Text = samples.ToString();
                }));
            }
            else
            {
                textBoxADCCh.Text = comboBoxADCCh.Items[ch].ToString();
                textBoxADCSps.Text = comboBoxSPS.Items[sps].ToString();
                textBoxADCSamples.Text = samples.ToString();
            }



        }

        private void Cmd1Parser(byte[] arr, int len)
        {
            byte cmd2 = arr[4];
            byte cmd3 = arr[5];
            byte[] data = new byte[len];
            Array.Copy(arr, 6, data, 0, len);

            switch (cmd2)
            {
                case 0x01: // LED_SET
                    switch (cmd3)
                    {
                        case 0x01:
                            break;
                        case 0x02:
                            LEDSettingParser(data);
                            break;
                    }
                    break;
                case 0x02:
                    switch (cmd3)
                    {
                        case 0x01:
                            break;
                        case 0x02:
                            ADCSettingParser(data);
                            break;
                    }
                    break;
            }
        }

        private void Cmd2Parser(byte[] arr, int len)
        {
            byte cmd2 = arr[4];
            byte cmd3 = arr[5];
            byte[] data = new byte[len];
            Array.Copy(arr, 6, data, 0, len);

            switch (cmd2)
            {
                case 0x00:
                    break;
            }
        }

        private void Cmd3Parser(byte[] arr, int len)
        {
            byte cmd2 = arr[4];
            byte cmd3 = arr[5];
            byte[] data = new byte[len - 6];
            Array.Copy(arr, 6, data, 0, len - 6);

            switch (cmd2)
            {
                case 0x00:
                    break;
                case 0x01:
                    switch(cmd3)
                    {
                        case 0x01:
                            ADCResultParser(data);
                            break;
                        case 0x02:
                            ADCMode3ResultParser(data);
                            break;
                    }
                    break;
                case 0x02:
                    
                    break;
            }
        }

        private void LEDSettingParser(byte[] data)
        {
        }

        private void ADCSettingFunc(byte ch, UInt16 sps, UInt16 adcNum)
        {
            byte[] bVal = new byte[7];
            UInt16 num = 200;

            bVal[0] = ch;
            bVal[1] = (byte)((sps >> 8) & 0xFF);
            bVal[2] = (byte)(sps & 0xFF);
            bVal[3] = (byte)((adcNum >> 8) & 0xFF);
            bVal[4] = (byte)(adcNum & 0xFF);
            bVal[3] = (byte)((num >> 8) & 0xFF);
            bVal[4] = (byte)(num & 0xFF);

            SendData(0x01, 0x02, 0x01, bVal, bVal.Length);
        }

        private void ADCSettingParser(byte[] data)
        {
            ADCSetting adc = new ADCSetting();
            adc.bVal = data;

            textBoxADCCh.Text = adc.CH == 0x00 ? "MONIT" : "RECEV";
            textBoxADCSps.Text = comboBoxSPS.Items[adc.SPS].ToString();
            textBoxADCSamples.Text = adc.SAMPLES.ToString();
        }

        private void ADCMode3ResultParser(byte[] data)
        {

        }

        private void ADCResultParser(byte[] data) // result
        {
            int chCnt = (data.Length) / 5;

            String textBoxName = "";

            switch (opMode)
            {
                case 0x01:
                    textBoxName = "textBoxMode1Ch";
                    break;
                case 0x02:
                    textBoxMode2Time.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    textBoxName = "textBoxMode2Ch";
                    break;
            }

            for (int i = 0; i < chCnt; i++)
            {
                BToFloat b2f;
                b2f.fVal = 0;

                int ch = data[i * 5];
                String targetName = textBoxName + ch + "Result";
                b2f.b4 = data[i * 5 + 4];
                b2f.b3 = data[i * 5 + 3];
                b2f.b2 = data[i * 5 + 2];
                b2f.b1 = data[i * 5 + 1];

                float value = b2f.fVal;

                var control = this.Controls.Find(targetName, true).FirstOrDefault();

                if (control != null)
                {
                    control.Text = value.ToString("0.000");

                    float min = -(float)((NumericUpDown)Controls.Find("numericUpDownPFCH" + ch + "Min", true).FirstOrDefault()).Value;
                    float max = -(float)((NumericUpDown)Controls.Find("numericUpDownPFCH" + ch + "Max", true).FirstOrDefault()).Value;
                    float err = -(float)((NumericUpDown)Controls.Find("numericUpDownPFCH" + ch + "Error", true).FirstOrDefault()).Value;

                    min = min - err;
                    max = max + err;

                    if (value >= min && value <= max)
                    {
                        // Pass
                        control.BackColor = Color.Green;
                    }
                    else
                    {
                        // Fail
                        control.BackColor = Color.Red;
                    }
                }
            }
        }

        private void SendData(byte cmd1, byte cmd2, byte cmd3, byte[] arr, int len)
        {
            if (sp == null || sp.IsConnected() == false)
            {
                MessageBox.Show("연결 후 눌러주세요.");
                return;
            }

            sp.SendData(cmd1, cmd2, cmd3, arr, len);
        }

        private void buttonLedOn_Click(object sender, EventArgs e)
        {
            // LED 켜기
            byte[] arr = new byte[4];

            arr[0] = (byte)((checkBoxCh1.Checked == true) ? 0x01 : 0x00);
            arr[1] = (byte)((checkBoxCh2.Checked == true) ? 0x01 : 0x00);
            arr[2] = (byte)((checkBoxCh3.Checked == true) ? 0x01 : 0x00);
            arr[3] = (byte)((checkBoxCh4.Checked == true) ? 0x01 : 0x00);

            SendData(0x03, 0x01, 0x00, arr, 4);
        }

        private void buttonLedOff_Click(object sender, EventArgs e)
        {
            // LED 끄기
            byte[] arr = new byte[4];

            SendData(0x03, 0x01, 0x00, arr, 4);
        }

        private void buttonPowerOn_Click(object sender, EventArgs e)
        {
            SendData(0x03, 0x01, 0x00, new byte[1] { 0x01 }, 1);
        }

        private void buttonPowerOff_Click(object sender, EventArgs e)
        {
            SendData(0x03, 0x01, 0x00, new byte[1] { 0x00 }, 1);
        }

        private void buttonIndRun_Click(object sender, EventArgs e)
        {
            seqList = new List<Sequence>();

            if (checkBoxCh1.Checked == true) // Blue
            {
                seqList.Add(new Sequence(0, (UInt16)numericUpDownCh0Setting.Value, 0.0f));
            }
            if (checkBoxCh2.Checked == true) // Green
            {
                seqList.Add(new Sequence(1, (UInt16)numericUpDownCh1Setting.Value, 0.0f));
            }
            if (checkBoxCh3.Checked == true) // Red
            {
                seqList.Add(new Sequence(2, (UInt16)numericUpDownCh2Setting.Value, 0.0f));
            }
            if (checkBoxCh4.Checked == true) // Orange
            {
                seqList.Add(new Sequence(3, (UInt16)numericUpDownCh3Setting.Value, 0.0f));
            }

            SendMeasSetting(seqList);

            InitMode1();
        }

        private void SendMeasSetting(List<Sequence> seqList)
        {
            int seqSize = seqList.Count();

            if (seqSize == 0)
            {
                MessageBox.Show("최소 하나의 채널은 테스트해야 합니다.");
                return;
            }

            byte[] bytes = new byte[seqSize * 3];

            for (int i = 0; i < seqSize; i++)
            {
                bytes[i * 3] = seqList[i].index;
                bytes[i * 3 + 1] = (byte)(seqList[i].outputValue >> 8);
                bytes[i * 3 + 2] = (byte)(seqList[i].outputValue & 0xFF);
            }

            SendData(0x01, 0x01, 0x01, bytes, bytes.Length);
        }

        private void buttonMeasurement_Click(object sender, EventArgs e)
        {
            opMode = 1;
            SendData(0x02, 0x01, 0x01, new byte[1] { 0x00 }, 1);

            InitMode1();
        }


        private void buttonMeasurementMode2_Click(object sender, EventArgs e)
        {
            opMode = 2;

            seqList = new List<Sequence>();

            seqList.Add(new Sequence(0, (UInt16)numericUpDownCh0Setting.Value, 0.0f));
            seqList.Add(new Sequence(1, (UInt16)numericUpDownCh1Setting.Value, 0.0f));
            seqList.Add(new Sequence(2, (UInt16)numericUpDownCh2Setting.Value, 0.0f));
            seqList.Add(new Sequence(3, (UInt16)numericUpDownCh3Setting.Value, 0.0f));

            SendMeasSetting(seqList);

            Thread.Sleep(10);

            SendData(0x02, 0x01, 0x01, new byte[1] { 0x00 }, 1);

            InitMode2();
        }

        private void buttonMode1Clear_Click(object sender, EventArgs e)
        {
            InitMode1();
        }

        private void InitMode1()
        {
            textBoxMode1Ch0Result.Text = "";
            textBoxMode1Ch1Result.Text = "";
            textBoxMode1Ch2Result.Text = "";
            textBoxMode1Ch3Result.Text = "";

            textBoxMode1Ch0Result.BackColor = SystemColors.Control;
            textBoxMode1Ch1Result.BackColor = SystemColors.Control;
            textBoxMode1Ch2Result.BackColor = SystemColors.Control;
            textBoxMode1Ch3Result.BackColor = SystemColors.Control;
        }

        private void buttonMode2Clear_Click(object sender, EventArgs e)
        {
            InitMode2();
        }

        private void InitMode2()
        {
            textBoxMode2Time.Text = "";
            textBoxMode2Ch0Result.Text = "";
            textBoxMode2Ch1Result.Text = "";
            textBoxMode2Ch2Result.Text = "";
            textBoxMode2Ch3Result.Text = "";

            textBoxMode2Ch0Result.BackColor = SystemColors.Control;
            textBoxMode2Ch1Result.BackColor = SystemColors.Control;
            textBoxMode2Ch2Result.BackColor = SystemColors.Control;
            textBoxMode2Ch3Result.BackColor = SystemColors.Control;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBoxMode3Ch1ResultMin.Text = "";
            textBoxMode3Ch1ResultMax.Text = "";
            textBoxMode3Ch2ResultMin.Text = "";
            textBoxMode3Ch2ResultMax.Text = "";
            textBoxMode3Ch3ResultMin.Text = "";
            textBoxMode3Ch3ResultMax.Text = "";
            textBoxMode3Ch4ResultMin.Text = "";
            textBoxMode3Ch4ResultMax.Text = "";

            textBoxMode3Ch1ResultMin.BackColor = SystemColors.Control;
            textBoxMode3Ch1ResultMax.BackColor = SystemColors.Control;
            textBoxMode3Ch2ResultMin.BackColor = SystemColors.Control;
            textBoxMode3Ch2ResultMax.BackColor = SystemColors.Control;
            textBoxMode3Ch3ResultMin.BackColor = SystemColors.Control;
            textBoxMode3Ch3ResultMax.BackColor = SystemColors.Control;
            textBoxMode3Ch4ResultMin.BackColor = SystemColors.Control;
            textBoxMode3Ch4ResultMax.BackColor = SystemColors.Control;
        }

        private void buttonSetDefault_Click(object sender, EventArgs e)
        {
            SetDefault();

            DisplayInit();
        }

        private void MeasureMode1()
        {

        }

        private void MeasureMode2()
        {

        }

        private void MeasureMode3()
        {
            SendData(0x02, 0x1, 0x02, new byte[1] { 0x00 }, 1);
        }

        private void buttonCh1Min_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch1ResultMin;
            MeasureMode3();
        }

        private void buttonCh1Max_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch1ResultMax;
            MeasureMode3();
        }

        private void buttonCh2Min_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch2ResultMin;
            MeasureMode3();
        }

        private void buttonCh2Max_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch2ResultMax;
            MeasureMode3();
        }

        private void buttonCh3Min_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch3ResultMin;
            MeasureMode3();
        }

        private void buttonCh3Max_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch3ResultMax;
            MeasureMode3();
        }

        private void buttonCh4Min_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch4ResultMin;
            MeasureMode3();
        }

        private void buttonCh4Max_Click(object sender, EventArgs e)
        {
            targetTextBox = textBoxMode3Ch4ResultMax;
            MeasureMode3();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseExcel();
        }

        private void CloseExcel()
        { }

        private void buttonDACSet_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.CH1_VALUE = (short)numericUpDownCh0Setting.Value;
            Properties.Settings.Default.CH2_VALUE = (short)numericUpDownCh1Setting.Value;
            Properties.Settings.Default.CH3_VALUE = (short)numericUpDownCh2Setting.Value;
            Properties.Settings.Default.CH4_VALUE = (short)numericUpDownCh3Setting.Value;

            Properties.Settings.Default.Save();
        }

        private void buttonADCSet_Click(object sender, EventArgs e)
        {
            byte chIdx = (byte)comboBoxADCCh.SelectedIndex;
            UInt16 spsIdx = (UInt16)comboBoxSPS.SelectedIndex;
            UInt16 samples = (UInt16)numericUpDownADCSamples.Value;

            ADCSettingFunc(chIdx, spsIdx, samples);
        }

        private void buttonTFInit_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.CH1_PF_MIN = 121;
            Properties.Settings.Default.CH1_PF_MAX = 189;
            Properties.Settings.Default.CH1_PF_ERROR = 0.5f;

            Properties.Settings.Default.CH2_PF_MIN = 28.8f;
            Properties.Settings.Default.CH2_PF_MAX = 33;
            Properties.Settings.Default.CH2_PF_ERROR = 0.5f;

            Properties.Settings.Default.CH3_PF_MIN = 174;
            Properties.Settings.Default.CH3_PF_MAX = 241.5f;
            Properties.Settings.Default.CH3_PF_ERROR = 0.5f;

            Properties.Settings.Default.CH4_PF_MIN = 61;
            Properties.Settings.Default.CH4_PF_MAX = 241.5f;
            Properties.Settings.Default.CH4_PF_ERROR = 0.5f;

            Properties.Settings.Default.Save();

            DisplayInit();
        }

        private void buttonTFSet_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.CH1_PF_MIN = (float)numericUpDownPFCH0Min.Value;
            Properties.Settings.Default.CH2_PF_MIN = (float)numericUpDownPFCH1Min.Value;
            Properties.Settings.Default.CH3_PF_MIN = (float)numericUpDownPFCH2Min.Value;
            Properties.Settings.Default.CH4_PF_MIN = (float)numericUpDownPFCH3Min.Value;

            Properties.Settings.Default.CH1_PF_MAX = (float)Properties.Settings.Default.CH1_PF_MAX;
            Properties.Settings.Default.CH2_PF_MAX = (float)Properties.Settings.Default.CH2_PF_MAX;
            Properties.Settings.Default.CH3_PF_MAX = (float)Properties.Settings.Default.CH3_PF_MAX;
            Properties.Settings.Default.CH4_PF_MAX = (float)Properties.Settings.Default.CH4_PF_MAX;

            Properties.Settings.Default.CH1_PF_ERROR = (float)numericUpDownPFCH0Error.Value;
            Properties.Settings.Default.CH2_PF_ERROR = (float)numericUpDownPFCH1Error.Value;
            Properties.Settings.Default.CH3_PF_ERROR = (float)numericUpDownPFCH2Error.Value;
            Properties.Settings.Default.CH4_PF_ERROR = (float)numericUpDownPFCH3Error.Value;

            Properties.Settings.Default.Save();
        }

        private void buttonWriteMode2Report_Click(object sender, EventArgs e)
        {
            if (textBoxSn.Text.Length == 0)
            {
                MessageBox.Show("시리얼 번호가 입력되어 있지 않습니다.");
                return;
            }

            if (textBoxMode2Ch0Result.Text.Length == 0)
            {
                MessageBox.Show("측정 결과가 없습니다.");
                return;
            }

            float[] value = new float[4];

            value[0] = float.Parse(textBoxMode2Ch0Result.Text);
            value[1] = float.Parse(textBoxMode2Ch1Result.Text);
            value[2] = float.Parse(textBoxMode2Ch2Result.Text);
            value[3] = float.Parse(textBoxMode2Ch3Result.Text);

            float[] minValue = new float[4];

            minValue[0] = (float)numericUpDownPFCH0Min.Value;
            minValue[1] = (float)numericUpDownPFCH1Min.Value;
            minValue[2] = (float)numericUpDownPFCH2Min.Value;
            minValue[3] = (float)numericUpDownPFCH3Min.Value;

            float[] maxValue = new float[4];

            maxValue[0] = (float)numericUpDownPFCH0Max.Value;
            maxValue[1] = (float)numericUpDownPFCH1Max.Value;
            maxValue[2] = (float)numericUpDownPFCH2Max.Value;
            maxValue[3] = (float)numericUpDownPFCH3Max.Value;

            float[] errValue = new float[4];
            errValue[0] = (float)numericUpDownPFCH0Error.Value;
            errValue[1] = (float)numericUpDownPFCH1Error.Value;
            errValue[2] = (float)numericUpDownPFCH2Error.Value;
            errValue[3] = (float)numericUpDownPFCH3Error.Value;

            WriteToExcelMode2(textBoxSn.Text, value, minValue, maxValue, errValue);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            WriteToExcelMode3();
        }

        private void buttonRescan_Click(object sender, EventArgs e)
        {
            GetPortName();
        }
    }
}
