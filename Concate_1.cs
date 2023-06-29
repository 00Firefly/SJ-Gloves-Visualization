using CsvHelper;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class Concate_1 : Form
    {
        #region 变量声明
        public int chan_num = 4;

        public const int mpu_num = 2;
        public const int ele_num = 7;
        public const int frames = 100;

        public byte[] uart_buf;
        public float[,,] raw_data;
        public float[] angle2chart;
        public float[,] FFF_chart;
        public float[] bias;
        public Queue<float> angle_queue;
        public Queue<float>[] raw_queue;
        public Queue<float>[] ad_queue;
        public Queue<int> marker_queue;
        public float[,] vec;
        public float[,] rpy_gyro;
        public float[,] rpy_acc;
        public float[,] rpy_kalman;


        public TextBox[] textBoxes;
        public Chart[] charts;
        public int[] line_nums;
        public string[,] line_names;
        public Color[] colors;

        public DateTime start_t, end_t;
        public TimeSpan ts;

        public int lastMove = 0;
        public bool isMouseDown = false;

        public int marker = 1;

        public bool isTriggered = false;

        float rad2deg = 57.29578f;
        //Kalman filter coefficients
        float[,] e_P = new float[2, 2] {
        { 1, 0},
        { 0, 1}
        };
        float[,] Q = new float[2, 2] {
        { 2.5e-3f, 0},
        { 0, 2.5e-3f}
        };
        float[,] R = new float[2, 2] {
        { 3e-1f, 0},
        { 0, 3e-1f}
        };
        float[,] k_k = new float[2, 2];

        private volatile bool is_serial_listening = false;//串口正在监听标记
        private volatile bool is_serial_closing = false;//串口正在关闭标记

        delegate void my_delegate();//创建一个代理,图表刷新需要在主线程，所以需要加委托
        #endregion
        public Concate_1()
        {
            InitializeComponent();
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
            serialPort1.Encoding = Encoding.GetEncoding("GB2312");
            pictureBox1.BringToFront();
            CheckForIllegalCrossThreadCalls = false;
        }
        private void Concate_1_Load(object sender, EventArgs e)
        {
            angle2chart = new float[frames];
            raw_data = new float[mpu_num, ele_num, frames];
            bias = new float[mpu_num * ele_num + chan_num];

            FFF_chart = new float[chan_num, frames];
            ad_queue = new Queue<float>[chan_num];
            marker_queue = new Queue<int>();
            rpy_gyro = new float[mpu_num, 3];
            rpy_acc = new float[mpu_num, 3];
            rpy_kalman = new float[mpu_num, 3];

            start_t = DateTime.Now;
            end_t = DateTime.Now;
            raw_queue = new Queue<float>[mpu_num * ele_num];
            for (int i = 0; i < mpu_num * ele_num; i++)
                raw_queue[i] = new Queue<float>();
            for (int i = 0; i < chan_num; i++)
                ad_queue[i] = new Queue<float>();
            angle_queue = new Queue<float>();

            textBoxes = new TextBox[] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7, textBox8,
                textBox9, textBox10, textBox11, textBox12, textBox13, textBox14, textBox15, textBox16, textBox17 };
            charts = new Chart[] { chart1, chart2, chart3, chart4, chart5, chart6 };
            line_nums = new int[] { 1, 4, 1, 1, 1, 1 };
            line_names = new string[,] {
                { "角度", "", "", "" },
                { "小指", "无名指", "中指", "食指" },
                { "小指", "", "", "" },
                { "无名指", "", "", "" },
                { "中指", "", "", "" },
                { "食指", "", "", "" }
            };
            colors = new Color[] { Color.Red, Color.Blue, Color.Green, Color.Purple };
            SearchAnAddSerialToComboBox(serialPort1, comboBox1);

            for (int i = 0; i < charts.Length; i++)
            {
                charts[i].ChartAreas[0].AxisX.Minimum = 0;
                charts[i].ChartAreas[0].AxisX.Maximum = frames;
                for (int j = 0; j < line_nums[i]; j++)
                {
                    charts[i].Series.Add(line_names[i, j]);
                    charts[i].Series[line_names[i, j]].Color = colors[j];
                    charts[i].Series[line_names[i, j]].ChartType = SeriesChartType.Line;
                }
            }
            foreach (TextBox textBox in textBoxes)
                textBox.Clear();
            pictureBox1.BackColor = Color.Red;
        }
        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (is_serial_closing)
            {
                is_serial_listening = false; //准备关闭串口时，reset串口侦听标记
                return;
            }
            try
            {
                if (serialPort1.IsOpen)
                {
                    is_serial_listening = true;
                    Draw();
                    uart_buf = new byte[serialPort1.BytesToRead];//定义缓冲数组
                    serialPort1.Read(uart_buf, 0, uart_buf.Length);
                    Chart_go(uart_buf);
                }
            }
            finally
            {
                is_serial_listening = false;//串口调用完毕后，reset串口侦听标记
            }
        }
        public void Draw()
        {
            try
            {
                if (!chart1.InvokeRequired)
                {
                    for (int i = 0; i < charts.Length; i++)
                    {
                        for (int j = 0; j < line_nums[i]; j++)
                        {
                            charts[i].Series[line_names[i, j]].Points.Clear();
                            for (int k = 0; k < frames; k++)
                            {
                                if (i == 0)
                                    charts[i].Series[line_names[i, j]].Points.AddXY(k, angle2chart[k]);
                                else if (i == 1)
                                    charts[i].Series[line_names[i, j]].Points.AddXY(k, FFF_chart[j, k]);
                                else
                                    charts[i].Series[line_names[i, j]].Points.AddXY(k, FFF_chart[i - 2, k]);
                            }
                        }
                    }
                }
                else
                {
                    my_delegate delegate_1 = new my_delegate(Draw);
                    Invoke(delegate_1, new object[] { });//执行唤醒操作
                }
            }
            catch { }
        }
        public void Chart_go(byte[] buf)
        {
            const int buf_size = 47;
            byte sum_check;
            float forceMapped;
            float[] temp = new float[mpu_num * ele_num + chan_num];
            float[] raw_data_slice = new float[mpu_num * ele_num + chan_num];
            byte[] buf_slice = new byte[buf_size];

            if ((buf.Length % buf_size == 0) && (buf.Length != 0) && (buf[0] == 255))
            {
                for (int j = 0; j < buf.Length / buf_size; j++)
                {
                    for (int i = 0; i < buf_size; i++)
                        buf_slice[i] = buf[i + buf_size * j];
                    sum_check = 0;
                    for (int i = 0; i < buf_slice.Length - 1; i++)
                        sum_check += buf_slice[i];
                    sum_check = (byte)~sum_check;
                    if (sum_check == buf_slice[buf_slice.Length - 1])
                    {
                        for (int i = 0; i < 2; i++)
                        {
                            for (int k = 0; k < ele_num - 1; k++)
                            {
                                temp[k + i * ele_num] = BitConverter.ToInt16(buf_slice, 1 + 2 * k + i * 18);
                                if (k < 3)
                                    temp[k + i * ele_num] /= 16384.0f;
                                else
                                    temp[k + i * ele_num] /= 131.0f;
                                temp[k + i * ele_num] -= bias[k + i * ele_num];
                            }
                            temp[ele_num - 1 + i * ele_num] = BitConverter.ToSingle(buf_slice, 13 + i * 18) / 340.0f + 36.53f - 8.0f;
                        }
                        for (int i = 0; i < chan_num; i++)
                            temp[mpu_num * ele_num + i] = (buf_slice[37 + 2 * i] << 8) + buf_slice[38 + 2 * i] - bias[mpu_num * ele_num + i];
                        //temp is ready
                        for (int i = 0; i < frames - 1; i++)
                        {
                            angle2chart[i] = angle2chart[i + 1];
                            for (int k = 0; k < ele_num * mpu_num; k++)
                                raw_data[k / ele_num, k % ele_num, i] = raw_data[k / ele_num, k % ele_num, i + 1];
                            for (int k = 0; k < chan_num; k++)
                                FFF_chart[k, i] = FFF_chart[k, i + 1];
                        }
                        for (int i = 0; i < mpu_num * ele_num; i++)
                        {
                            raw_queue[i].Enqueue(temp[i]);
                            raw_data[i / ele_num, i % ele_num, frames - 1] = temp[i];
                        }
                        for (int i = 0; i < mpu_num * ele_num; i++)
                        {
                            raw_data_slice[i] = raw_data[i / ele_num, i % ele_num, frames - 1];
                            if (((i >= 3) && (i < 6)) || ((i >= 10) && (i < 13)))
                                raw_data_slice[i] /= rad2deg;
                        }
                        end_t = DateTime.Now;
                        ts = end_t - start_t;
                        angle2chart[frames - 1] = AngleCalc(raw_data_slice, (float)ts.TotalSeconds);
                        start_t = DateTime.Now;
                        angle_queue.Enqueue(angle2chart[frames - 1]);
                        for (int i = 0; i < chan_num; i++)
                        {
                            forceMapped = ForceMap(temp[mpu_num * ele_num + i], i);
                            ad_queue[i].Enqueue(forceMapped);

                            if (i == 3 && forceMapped > 0.5 && !isTriggered)
                            {
                                marker = 2;
                                isTriggered = true;
                            }
                            else if (i == 3 && forceMapped < 0.4 && isTriggered)
                            {
                                marker = 3;
                                isTriggered = false;
                            }

                            if (checkBox5.Checked)
                                FFF_chart[i, frames - 1] = forceMapped;
                            else
                                FFF_chart[i, frames - 1] = temp[mpu_num * ele_num + i];
                        }
                        marker_queue.Enqueue(marker);
                        Text_go();
                    }
                }
            }
        }
        public void Text_go()
        {
            textBox2.Text = raw_data[0, 1, frames - 1].ToString("f2");
            textBox3.Text = raw_data[0, 2, frames - 1].ToString("f2");
            textBox4.Text = raw_data[0, 3, frames - 1].ToString("f2");
            textBox5.Text = raw_data[0, 4, frames - 1].ToString("f2");
            textBox6.Text = raw_data[0, 5, frames - 1].ToString("f2");
            textBox7.Text = raw_data[1, 0, frames - 1].ToString("f2");
            textBox8.Text = raw_data[1, 1, frames - 1].ToString("f2");
            textBox9.Text = raw_data[1, 2, frames - 1].ToString("f2");
            textBox10.Text = raw_data[1, 3, frames - 1].ToString("f2");
            textBox11.Text = raw_data[1, 4, frames - 1].ToString("f2");
            textBox12.Text = raw_data[1, 5, frames - 1].ToString("f2");
            //textBox1.DataBindings.Add("Text", raw_data[0, 0, frames - 1], "Value", true, DataSourceUpdateMode.OnPropertyChanged, 0, "N2");
            if (!checkBox5.Checked)
            {
                textBox14.Text = FFF_chart[0, frames - 1].ToString("f2");
                textBox15.Text = FFF_chart[1, frames - 1].ToString("f2");
                textBox16.Text = FFF_chart[2, frames - 1].ToString("f2");
                textBox17.Text = FFF_chart[3, frames - 1].ToString("f2");
            }
            else
            {
                textBox14.Text = FFF_chart[0, frames - 1].ToString("f2");
                textBox15.Text = FFF_chart[1, frames - 1].ToString("f2");
                textBox16.Text = FFF_chart[2, frames - 1].ToString("f2");
                textBox17.Text = FFF_chart[3, frames - 1].ToString("f2");
            }
        }
        private float AngleCalc(float[] mpu, float ts_s)
        {
            float angle;
            float[,] trans_matrix_gyro;
            float[,] vec;
            float temp;

            for (int k = 0; k < mpu_num; k++)
            {
                float[] ryp_gyro_slice = new float[] { rpy_gyro[k, 0], rpy_gyro[k, 1], rpy_gyro[k, 2] };
                trans_matrix_gyro = TransMatrixGyro(ryp_gyro_slice);

                for (int i = 0; i < 3; i++)
                {
                    temp = 0;
                    for (int j = 0; j < 3; j++)
                        temp += trans_matrix_gyro[i, j] * mpu[3 + j + k * 7];
                    rpy_gyro[k, i] += temp * ts_s;
                }

                for (int i = 0; i < 2; i++)
                    for (int j = 0; j < 2; j++)
                        e_P[i, j] += Q[i, j];
                for (int i = 0; i < 2; i++)
                    for (int j = 0; j < 2; j++)
                        if (e_P[i, j] != 0)
                            k_k[i, j] += e_P[i, j] / (e_P[i, j] + R[i, j]);

                rpy_acc[k, 0] = (float)Math.Atan2(mpu[1 + k * 7], mpu[2 + k * 7]);
                rpy_acc[k, 1] = (float)-Math.Atan2(mpu[k * 7], Math.Sqrt(Math.Pow(mpu[1 + k * 7], 2) + Math.Pow(mpu[2 + k * 7], 2)));

                rpy_kalman[k, 0] = rpy_gyro[k, 0] + k_k[0, 0] * (rpy_acc[k, 0] - rpy_gyro[k, 0]);
                rpy_kalman[k, 1] = rpy_gyro[k, 1] + k_k[1, 1] * (rpy_acc[k, 1] - rpy_gyro[k, 1]);
                rpy_kalman[k, 2] = rpy_gyro[k, 2];

                for (int i = 0; i < 2; i++)
                    for (int j = 0; j < 2; j++)
                        e_P[i, j] -= e_P[i, j] * k_k[i, j];
            }

            //vec = new float[mpu_num, 3] {
            //    { (float)(Math.Cos(rpy_kalman[0, 1]) * Math.Cos(rpy_kalman[0, 2])), (float)(Math.Sin(rpy_kalman[0, 0]) * Math.Sin(rpy_kalman[0, 1]) * Math.Cos(rpy_kalman[0, 2]) - Math.Cos(rpy_kalman[0, 0]) * Math.Sin(rpy_kalman[0, 2])), (float)(Math.Cos(rpy_kalman[0, 0]) * Math.Sin(rpy_kalman[0, 1]) * Math.Cos(rpy_kalman[0, 2]) + Math.Sin(rpy_kalman[0, 0]) * Math.Sin(rpy_kalman[0, 2])) },
            //    { (float)(Math.Cos(rpy_kalman[1, 1]) * Math.Cos(rpy_kalman[1, 2])), (float)(Math.Sin(rpy_kalman[1, 0]) * Math.Sin(rpy_kalman[1, 1]) * Math.Cos(rpy_kalman[1, 2]) - Math.Cos(rpy_kalman[1, 0]) * Math.Sin(rpy_kalman[1, 2])), (float)(Math.Cos(rpy_kalman[1, 0]) * Math.Sin(rpy_kalman[1, 1]) * Math.Cos(rpy_kalman[1, 2]) + Math.Sin(rpy_kalman[1, 0]) * Math.Sin(rpy_kalman[1, 2])) }
            //};
            vec = new float[mpu_num, 3] {
                { (float)(Math.Cos(rpy_acc[0, 1]) * Math.Cos(rpy_gyro[0, 2])), (float)(Math.Sin(rpy_acc[0, 0]) * Math.Sin(rpy_acc[0, 1]) * Math.Cos(rpy_gyro[0, 2]) - Math.Cos(rpy_acc[0, 0]) * Math.Sin(rpy_gyro[0, 2])), (float)(Math.Cos(rpy_acc[0, 0]) * Math.Sin(rpy_acc[0, 1]) * Math.Cos(rpy_gyro[0, 2]) + Math.Sin(rpy_acc[0, 0]) * Math.Sin(rpy_gyro[0, 2])) },
                { (float)(Math.Cos(rpy_acc[1, 1]) * Math.Cos(rpy_gyro[1, 2])), (float)(Math.Sin(rpy_acc[1, 0]) * Math.Sin(rpy_acc[1, 1]) * Math.Cos(rpy_gyro[1, 2]) - Math.Cos(rpy_acc[1, 0]) * Math.Sin(rpy_gyro[1, 2])), (float)(Math.Cos(rpy_acc[1, 0]) * Math.Sin(rpy_acc[1, 1]) * Math.Cos(rpy_gyro[1, 2]) + Math.Sin(rpy_acc[1, 0]) * Math.Sin(rpy_gyro[1, 2])) }
            };

            angle = (float)Math.Acos((vec[0, 0] * vec[1, 0] + vec[0, 1] * vec[1, 1] + vec[0, 2] * vec[1, 2])
                / Math.Sqrt((vec[0, 0] * vec[0, 0] + vec[0, 1] * vec[0, 1] + vec[0, 2] * vec[0, 2])
                * (vec[1, 0] * vec[1, 0] + vec[1, 1] * vec[1, 1] + vec[1, 2] * vec[1, 2])));
            angle *= rad2deg;
            return angle;
        }
        private float[,] TransMatrixGyro(float[] ryp)
        {
            float r = ryp[0];
            float y = ryp[1];
            float p = ryp[2];
            float[,] m = new float[3, 3] {
            { 1, (float)(Math.Tan(p) * Math.Sin(r)), (float)(Math.Tan(p) * Math.Cos(r))},
            { 0, (float)Math.Cos(r), (float)-Math.Sin(r)},
            { 0, (float)(Math.Sin(r) / Math.Cos(p)), (float)(Math.Cos(r) / Math.Cos(p))}
            };
            return m;
        }
        public float ForceMap(float ad, int finger_index)
        {
            float[] force_list = new float[28] {
                0.2f, 0.3f, 0.4f, 0.5f, 0.6f, 0.7f, 0.8f, 0.9f, 1, 1.5f, 2, 2.5f,
                3, 3.5f, 4, 4.5f, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
            };
            float[,] ad_lists = new float[4, 28]
            {
                { 1700,2770,3120,3271,3368,3518,3573,3613,3645,3742,3791,3817,3835,3849,3860,3867,3872,3881,3887,3891,3895,3898,3902,3904,3907,3909,3910,3910.5f},
                { 1120,2398,3100,3264,3439,3562,3630,3661,3688,3755,3797,3822,3839,3855,3866,3873,3878,3887,3893,3896,3902,3906,3909,3911,3913,3914,3916,3917},
                { 1130,2363,2724,2982,3111,3303,3414,3442,3484,3623,3685,3722,3746,3763,3773,3787,3802,3822,3830,3833,3834,3837,3838,3838.5f,3839,3839.5f,3840,3840.5f},
                { 1200,2136,2435,2726,2843,2978,3150,3290,3339,3539,3607,3701,3758,3794,3819,3840,3851,3872,3881,3888,3893,3898,3902,3904,3906,3908,3910,3911 }
            };
            float force;
            int index_ad = 0;
            if (ad < ad_lists[finger_index, 0])
                return 0;
            else
            {
                while (ad_lists[finger_index, index_ad] < ad)
                {
                    if (index_ad == 27)
                        return force_list[force_list.Length - 1];
                    else
                        index_ad++;
                }
                force = (ad - ad_lists[finger_index, index_ad - 1]) / (ad_lists[finger_index, index_ad] - ad_lists[finger_index, index_ad - 1])
                    * (force_list[index_ad] - force_list[index_ad - 1]) + force_list[index_ad - 1];
                return force;
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            SearchAnAddSerialToComboBox(serialPort1, comboBox1);
        }
        private void SearchAnAddSerialToComboBox(SerialPort MyPort, ComboBox MyBox)//搜索串口函数
        { //将可用的串口号添加到ComboBox
            string[] NmberOfport = new string[10];//最多容纳20个，太多会卡，影响效率
            string MidString1;//中间数组，用于缓存
            MyBox.Items.Clear();//清空combobox的内容
            for (int i = 1; i < 10; i++)
            {
                try //核心是靠try和catch 完成遍历
                {
                    MidString1 = "COM" + i.ToString();  //把串口名字赋给MidString1
                    MyPort.PortName = MidString1;       //把MidString1赋给 MyPort.PortName 
                    MyPort.Open();                      //如果失败，后面代码不执行？？
                    NmberOfport[i - 1] = MidString1;    //依次把MidString1的字符赋给NmberOfport
                    MyBox.Items.Add(MidString1);        //打开成功，添加到下列列表
                    MyPort.Close();                     //关闭
                    MyBox.Text = NmberOfport[i - 1];    //显示最后扫描成功那个串口
                }
                catch { };
            }
        }
        private void Commander(int marker)
        {
            byte[] command_1 = new byte[] { 0xFF, 0x00, 0x00, (byte)marker, 0xFF };
            if (serialPort1.IsOpen)
            {
                serialPort1.Write(command_1, 0, command_1.Length);
            }
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            LightBulb_Paint(marker);
            Commander(marker);
            marker = 1;
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.PortName = comboBox1.Text;
                    serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                    serialPort1.Open();
                    button2.Text = " 关闭串口";
                    timer1.Enabled = true;
                    is_serial_closing = false;
                }
                catch
                {
                    MessageBox.Show(e.ToString());
                }
            }
            else
            {
                try
                {
                    timer1.Enabled = false;
                    is_serial_closing = true;
                    while (is_serial_listening)
                        Application.DoEvents();
                    serialPort1.Close();
                    button2.Text = "打开串口";
                }
                catch
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            foreach (TextBox textBox in textBoxes)
                textBox.Clear();
            foreach (Chart chart in charts)
                foreach (Series series in chart.Series)
                    series.Points.Clear();
            angle_queue.Clear();
            for (int i = 0; i < mpu_num * ele_num; i++)
                raw_queue[i].Clear();
            for (int i = 0; i < chan_num; i++)
                ad_queue[i].Clear();
            marker_queue.Clear();
        }
        private string QueueIndex2String(Queue<float>[] ad_queue)
        {
            string temp_string = "";
            int last_index;
            for (int i = 0; i < chan_num; i++)
            {
                temp_string += ad_queue[i].Dequeue().ToString("f2") + ",";
            }
            //去掉最后多余的逗号
            last_index = temp_string.LastIndexOf(',');
            if (last_index != -1)
                temp_string = temp_string.Substring(0, last_index);
            return temp_string;
        }
        private string Queues2String(Queue<float>[] q_temp)
        {
            string temp_string = "";
            int last_index;
            for (int i = 0; i < mpu_num * ele_num; i++)
            {
                temp_string += q_temp[i].Dequeue().ToString("f2") + ",";
            }
            //去掉最后多余的逗号
            last_index = temp_string.LastIndexOf(',');
            if (last_index != -1)
                temp_string = temp_string.Substring(0, last_index);
            return temp_string;
        }
        private void Button4_Click(object sender, EventArgs e)
        {
            int marker_temp;
            //保存3种数据：原始MPU6050数据、角度数据和力数据
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
                button2.Text = "打开串口";
            }
            string date_time = DateTime.Now.ToString("yyyyMMddHHmm");
            SaveFileDialog saveCsvDialog = new SaveFileDialog
            {
                Filter = "(*.csv)|*.csv|(*.*)|*.*",
                Title = "选择文件保存路径",
                RestoreDirectory = true,
            };
            //原始数据保存
            try
            {
                saveCsvDialog.FileName = "gyro_raw" + date_time + ".csv";
                if (saveCsvDialog.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter streamWriter = new StreamWriter(saveCsvDialog.FileName, true);
                    int last_index;
                    string header = "";
                    string[] headers = { "ax1", "ay1", "az1", "gx1", "gy1", "gz1", "temper1",
                        "ax2", "ay2", "az2", "gx2", "gy2", "gz2", "temper2"};
                    for (int i = 0; i < mpu_num * ele_num; i++)
                        header += headers[i] + ",";
                    header += "marker,";
                    last_index = header.LastIndexOf(',');
                    if (last_index != -1)
                        header = header.Substring(0, last_index);
                    streamWriter.WriteLine(header);
                    while (raw_queue[0].Count > 0)
                    {
                        marker_temp = marker_queue.Dequeue();
                        streamWriter.WriteLine(Queues2String(raw_queue) + "," + marker_temp.ToString());
                        marker_queue.Enqueue(marker_temp);
                    }
                    streamWriter.Close();
                    MessageBox.Show("原始数据保存成功");
                }
            }
            catch
            {
                MessageBox.Show("原始数据保存失败");
            }
            //角度数据保存
            try
            {
                saveCsvDialog.FileName = "gyro_angle" + date_time + ".csv";
                if (saveCsvDialog.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter streamWriter = new StreamWriter(saveCsvDialog.FileName, true);
                    int last_index;
                    string header = "";
                    header += "角度" + "," + "marker" + ",";
                    last_index = header.LastIndexOf(',');
                    if (last_index != -1)
                        header = header.Substring(0, last_index);
                    streamWriter.WriteLine(header);
                    while (angle_queue.Count > 0)
                    {
                        marker_temp = marker_queue.Dequeue();
                        streamWriter.WriteLine(angle_queue.Dequeue().ToString("f2") + "," + marker_temp.ToString());
                        marker_queue.Enqueue(marker_temp);
                    }
                    streamWriter.Close();
                    MessageBox.Show("角度数据保存成功");
                }
            }
            catch
            {
                MessageBox.Show("角度数据保存失败");
            }
            try
            {
                saveCsvDialog.FileName = "force" + date_time + ".csv";
                if (saveCsvDialog.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter streamWriter = new StreamWriter(saveCsvDialog.FileName, true);
                    int last_index;
                    string header = "";
                    for (int i = 0; i < chan_num; i++)
                        header += line_names[1, i] + ",";
                    header += "marker,";
                    last_index = header.LastIndexOf(',');
                    if (last_index != -1)
                        header = header.Substring(0, last_index);
                    streamWriter.WriteLine(header);
                    while (ad_queue[0].Count > 0)
                    {
                        marker_temp = marker_queue.Dequeue();
                        streamWriter.WriteLine(QueueIndex2String(ad_queue) + "," + marker_temp.ToString());
                        marker_queue.Enqueue(marker_temp);
                    }
                    streamWriter.Close();
                    MessageBox.Show("力数据保存成功");
                }
            }
            catch
            {
                MessageBox.Show("力数据保存失败");
            }
        }
        private void Button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "(*.csv)|*.csv|(*.*)|*.*"
            };
            if (serialPort1.IsOpen)
                MessageBox.Show("请先关闭串口");
            else if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetFileNameWithoutExtension(openFileDialog.FileName).StartsWith("gyro_angle"))
                {
                    chart1.Series["角度"].Points.Clear();

                    List<Dictionary<string, double>> records = new List<Dictionary<string, double>>();
                    List<string> columnNames = new List<string>();

                    using (var reader = new StreamReader(openFileDialog.FileName))
                    using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.CurrentCulture))
                    {
                        csv.Read();
                        csv.ReadHeader();
                        columnNames = csv.HeaderRecord.ToList();
                        while (csv.Read())
                        {
                            var record = new Dictionary<string, double>();
                            foreach (var column in columnNames)
                            {
                                record[column] = csv.GetField<double>(column);
                            }

                            records.Add(record);
                        }
                    }
                    chart1.ChartAreas[0].AxisX.Maximum = double.NaN;
                    foreach (var record in records)
                    {
                        chart1.Series["角度"].Points.AddY(record[columnNames[0]]);
                    }
                }
                else if (Path.GetFileNameWithoutExtension(openFileDialog.FileName).StartsWith("force"))
                {
                    for (int i = 1; i < charts.Length; i++)
                        for (int j = 0; j < line_nums[i]; j++)
                            charts[i].Series[line_names[i, j]].Points.Clear();

                    List<Dictionary<string, double>> records = new List<Dictionary<string, double>>();
                    List<string> columnNames = new List<string>();

                    using (var reader = new StreamReader(openFileDialog.FileName))
                    using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.CurrentCulture))
                    {
                        csv.Read();
                        csv.ReadHeader();
                        columnNames = csv.HeaderRecord.ToList();
                        while (csv.Read())
                        {
                            var record = new Dictionary<string, double>();
                            foreach (var column in columnNames)
                            {
                                record[column] = csv.GetField<double>(column);
                            }

                            records.Add(record);
                        }
                    }
                    chart1.ChartAreas[0].AxisX.Maximum = double.NaN;
                    foreach (var record in records)
                    {
                        for (int i = 0; i < chan_num; i++)
                        {
                            chart2.Series[line_names[1, i]].Points.AddY(record[columnNames[i]]);
                            charts[i + 2].Series[line_names[i + 2, 0]].Points.AddY(record[columnNames[i]]);
                        }
                    }
                }
                else
                    MessageBox.Show("无法打开文件。非法的文件名：" + Path.GetFileNameWithoutExtension(openFileDialog.FileName));
            }
        }
        public void BiasInit(float[,,] raw2bias, float[,] force2bias)
        {
            float[] bias_temp = new float[mpu_num * ele_num + chan_num];
            for (int i = 0; i < mpu_num * ele_num + chan_num; i++)
                bias_temp[i] = bias[i];
            bias = new float[mpu_num * ele_num + chan_num];

            for (int i = 0; i < mpu_num * ele_num + chan_num; i++)
            {
                for (int j = 0; j < frames; j++)
                    if (((i >= 10) && (i <= 12)) || ((i >= 3) && (i <= 5)))
                        bias[i] += raw2bias[i / ele_num, i % ele_num, j];
                    else if (i >= mpu_num * ele_num)
                        bias[i] += force2bias[i - mpu_num * ele_num, j];
                bias[i] /= frames;
                bias[i] += bias_temp[i];
            }
            rpy_gyro = new float[mpu_num, 3];
            rpy_acc = new float[mpu_num, 3];
            e_P = new float[2, 2] {
            { 1, 0},
            { 0, 1}
            };
            Q = new float[2, 2] {
            { 2.5e-3f, 0},
            { 0, 2.5e-3f}
            };
            R = new float[2, 2] {
            { 3e-1f, 0},
            { 0, 3e-1f}
            };
            k_k = new float[2, 2];
        }
        private async void Button9_ClickAsync(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
                MessageBox.Show("请打开串口");
            else
            {
                var msgBox = new InitWait("正在初始化...");
                msgBox.Show();
                await Task.Delay(TimeSpan.FromSeconds(1.5));
                BiasInit(raw_data, FFF_chart);
                msgBox.Close();
            }
        }
        private void LightBulb_Paint(int marker)
        {
            if (marker == 2)
            {
                pictureBox1.BackColor = Color.Green;
            }
            else if (marker == 3)
            {
                pictureBox1.BackColor = Color.Red;
            }
        }
        #region Mouse Operation
        private void Chart_MouseWheel(object sender, MouseEventArgs e)
        {
            Chart chart = (Chart)sender;
            if (double.IsNaN(chart.ChartAreas[0].AxisX.ScaleView.Size))
                if (serialPort1.IsOpen)
                    chart.ChartAreas[0].AxisX.ScaleView.Size = 100;
                else
                    chart.ChartAreas[0].AxisX.ScaleView.Size = chart.ChartAreas[0].AxisX.Maximum;
            if (chart.ChartAreas[0].AxisX.ScaleView.Size > 0)
            {
                chart.ChartAreas[0].AxisX.ScaleView.Size += e.Delta / 12 * 10;
            }
            else if (e.Delta > 0)
            {
                chart.ChartAreas[0].AxisX.ScaleView.Size += e.Delta / 12 * 10;
            }
        }
        private void Chart_MouseDown(object sender, MouseEventArgs e)
        {
            lastMove = e.X;
            isMouseDown = true;
        }
        private void Chart_MouseUp(object sender, MouseEventArgs e)
        {
            isMouseDown = false;
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 4; i++)
                charts[2 + i].Visible = false;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            Invalidate();
        }

        private void Button8_Click(object sender, EventArgs e)
        {

        }

        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            Chart chart = (Chart)sender;
            if (isMouseDown)
            {
                chart.ChartAreas[0].AxisX.ScaleView.Position += (int)((double)(lastMove - e.X) / chart.Size.Width * chart.ChartAreas[0].AxisX.ScaleView.Size);
                lastMove = e.X;
            }
        }
        #endregion
        private void Concate_1_FormClosing(object sender, FormClosingEventArgs e)
        {
            is_serial_closing = true;//关闭窗口时，置位is_serial_closing标记
            while (is_serial_listening) Application.DoEvents();
            serialPort1.Close();
        }
    }
}
