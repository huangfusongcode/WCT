/********************************************

// 文件名：OC_Project

// 文件功能描述：WCT 功能（完成Panel亮度色坐标测量部分，色坐标校正部分还需完善）

// 创建人：黄福松

// 创建时间：2019/11/12

// 描述：

// 修改人：

// 修改时间：

// 修改描述：

//******************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using HID;
using System.Text.RegularExpressions;
using System.Text;
using System.Drawing;

namespace OC_Porject_20191112
{
    public partial class Form1 : Form
    {
        //创建全局变量，用来存放已show出的窗体对象
        List<Form> allforms = new List<Form>();
        List<Thread> threadList = new List<Thread>();

        public static int Gray_Max = 255;
        public static string Tar_Lum;
        public static string Tar_x;
        public static string Tar_y;

        public static string Cur_Lum_Text;
        public static string Cur_X_Text;
        public static string Cur_Y_Text;
        public static string Gray_Pattern_Num_Str;

        public static string Std_Lum;
        public static string Std_X;
        public static string Std_Y;

        static int Read_Bin_File_Num = 4;
        byte[] Read_G2V_LUT = new byte[16384 * Read_Bin_File_Num];

        bool result;
        string TestGray;

        //建立表格存储数据
        DataTable Dt_Cur_Lum = new DataTable();
        DataTable Dt_Cur_X = new DataTable();
        DataTable Dt_Cur_Y = new DataTable();

        //开启新线程
        Thread Th_Initial_Pra;                                                     //初始化参数线程
        Thread Th_G2V_Correct;                                                     //色坐标校正线程
        public Thread Th_Tran_RGBPattern_To_Panel;                                 //Pattern传输线程
        Thread Th_FormAutosize;                                                    //窗体大小自适应线程

        //创建新对象实例
        Auto_Size_Suitble asc = new Auto_Size_Suitble();                           //控件随窗体变化自适应
        WCT_RGB wct_rgb = new WCT_RGB();                                           //WCT RGB
        CA410_Measurement ca410_measurement = new CA410_Measurement();             //410对象
        USB usb = new USB();                                                       //USB对象
        DigitalGamma digitalgamma = new DigitalGamma();                            //D_GMA
        CLA cla = new CLA();                                                       //CLA
        WCT wct = new WCT();                                                       //WCT RGBW

        const double mx = 0.30;
        const double my = 0.31;
        const double ml = 100;
        const double gma = 2.2;
        byte[] para = new byte[60];
        //----------------------------HID---------------------------------------------------------------------------------------------------------------------------------------------
        //定义Register
        byte DeviceAdder = 63;
        //初始化vendor ID和productor ID
        private UInt16 vendorId = 0x8888;
        private UInt16 productorId = 0xFFFF;


        //初始化HID类的对象:hid_usb
        Hid hidUsb = new Hid();

        //声明变量和初始化设备查找结果
        Hid.HID_RETURN hidReturn = Hid.HID_RETURN.NO_DEVICE_CONECTED;

        //声明一个报文类的对象
        report hid_report;

        // 写Buffer，从上位机输出到下位机
        public byte[] wbuff = new byte[64];
        public byte[] rbuff = new byte[64];

        public EventHandler Gray_Num_Change;
        /// <summary>
        /// Form1
        /// </summary>
        /// </param>
        public Form1()
        {
            InitializeComponent();
            TextBox.CheckForIllegalCrossThreadCalls = false;
            nUD_gamma.Text = gma.ToString("f1");
            nUD_tarlum.Text = ml.ToString();
            nUD_mx.Text = mx.ToString("f2");
            nUD_my.Text = my.ToString("f2");

            rBtn_RGB.Checked = true;
            rBtn_adr8.Checked = true;
            rBtn_data10.Checked = true;
            rbtn_G2V_adr10.Checked = true;
            rbtn_G2V_data12.Checked = true;

            usb.InitCompo();
            UsbDeviceHandle();
            UsbStateUpdate();

            Dt_Cur_Lum.Rows.Add();
            Dt_Cur_Lum.Columns.Add();
            Dt_Cur_X.Rows.Add();
            Dt_Cur_X.Columns.Add();
            Dt_Cur_Y.Rows.Add();
            Dt_Cur_Y.Columns.Add();

            wct.dt_rgbcomp.Columns.Add("Point", typeof(String));
            wct.dt_rgbcomp.Columns.Add("Lum", typeof(String));
            wct.dt_rgbcomp.Columns.Add("α", typeof(String));
            wct.dt_rgbcomp.Columns.Add("β", typeof(String));
            wct.dt_rgbcomp.Columns.Add("γ", typeof(String));

            wct_rgb.wct_dt.Columns.Add("Point", typeof(String));
            wct_rgb.wct_dt.Columns.Add("Num", typeof(String));
            wct_rgb.wct_dt.Columns.Add("R", typeof(String));
            wct_rgb.wct_dt.Columns.Add("G", typeof(String));
            wct_rgb.wct_dt.Columns.Add("B", typeof(String));
            wct_rgb.wct_dt.Columns.Add("W", typeof(String));

            wct_rgb.G2V_410MeasData.Columns.Add("Gray", typeof(String));
            wct_rgb.G2V_410MeasData.Columns.Add("Gray_Num", typeof(Int16));
            wct_rgb.G2V_410MeasData.Columns.Add("Lum", typeof(Double));
            wct_rgb.G2V_410MeasData.Columns.Add("X", typeof(Double));
            wct_rgb.G2V_410MeasData.Columns.Add("Y", typeof(Double));
            //wct_rgb.G2V_410MeasData.Columns.Add();
            //wct_rgb.G2V_410MeasData.Columns.Add();
            //wct_rgb.G2V_410MeasData.Columns.Add();
            //wct_rgb.G2V_410MeasData.Columns.Add();
            //wct_rgb.G2V_410MeasData.Columns.Add();

            //Datatable存数据同步到UI
            wct_rgb.G2V_Correct_Num_Tab.Rows.Add();
            wct_rgb.G2V_Correct_Num_Tab.Columns.Add();
            wct_rgb.G2V_Trans_Num_Tab.Rows.Add();
            wct_rgb.G2V_Trans_Num_Tab.Columns.Add();
            wct_rgb.G2V_Current_Lum.Rows.Add();
            wct_rgb.G2V_Current_Lum.Columns.Add();
            wct_rgb.G2V_Current_X.Rows.Add();
            wct_rgb.G2V_Current_X.Columns.Add();
            wct_rgb.G2V_Current_Y.Rows.Add();
            wct_rgb.G2V_Current_Y.Columns.Add();
            wct_rgb.G2V_Max_Lum_Limit.Rows.Add();
            wct_rgb.G2V_Max_Lum_Limit.Columns.Add();


            for (int i = 0; i < 4096; i++)
            {
                wct.dt_rgbcomp.Rows.Add();
            }
            for (int i = 0; i < 4096; i++)
            {
                wct_rgb.wct_dt.Rows.Add();

            }
            for (int i = 0; i < 3072; i++)
            {
                wct_rgb.G2V_410MeasData.Rows.Add();
            }
            //导出到EXCEL数值会有单引号，暂用表格值初始化方法解决
            for (int i = 0; i < 3072; i++)
            {
                if(i < 1024)
                {
                    wct_rgb.G2V_410MeasData.Rows[i][0] = "R";
                    wct_rgb.G2V_410MeasData.Rows[i][1] = i;
                }
                else if(i < 2048)
                {
                    wct_rgb.G2V_410MeasData.Rows[i][0] = "G";
                    wct_rgb.G2V_410MeasData.Rows[i][1] = i - 1024;
                }
                else
                {
                    wct_rgb.G2V_410MeasData.Rows[i][0] = "B";
                    wct_rgb.G2V_410MeasData.Rows[i][1] = i - 2048;
                }
                wct_rgb.G2V_410MeasData.Rows[i][2] = 0;
                wct_rgb.G2V_410MeasData.Rows[i][3] = 0;
                wct_rgb.G2V_410MeasData.Rows[i][4] = 0;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            Th_Initial_Pra = new Thread(Initial_Prameter);
            Th_Initial_Pra.IsBackground = true;
            Th_Initial_Pra.Start();
        }

        /// <summary>
        /// 自适应窗体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            Th_FormAutosize = new Thread(FormAutoSize);
            Th_FormAutosize.IsBackground = true;
            Th_FormAutosize.Start();

            asc.controlAutoSize(this);                         //窗体控件大小自适应，控件适应较慢，先屏蔽掉

        }
        private void FormAutoSize()
        {
            MessageBox.Show("窗体自适应");
        }
        /// <summary>
        /// 初始化参
        /// 数
        /// </summary>
        /// </param>
        /// 
        private void Initial_Prameter()
        {
            try
            {
                UsbStateUpdate();
                asc.controllInitializeSize(this);

                Targ_Lum_Text.Text = "150";
                Targ_X_Text.Text = "0.30";
                Targ_Y_Text.Text = "0.31";
                Gray_Pattern_Num_Str = "1024";

                wct_rgb.Std_Lum = (double)nUpDown_Std_Lum.Value;
                wct_rgb.Std_X = (double)nUpDown_Std_X.Value;
                wct_rgb.Std_Y = (double)nUpDown_Std_Y.Value;
                wct_rgb.G2V_LUT_Correction_Num = 0;
                wct_rgb.G2V_Gray_Pattern_Num = Convert.ToInt16(Compbox_Gray_Pattern_Num.Text);    //选择需要传输的灰阶总数，目前1024可用，其余待开发
                wct_rgb.Delay = Convert.ToInt16(comboBox_MeasInterval.Text);                      //量测间隔时间，单位ms


                CA410_Measurement.chnum = 0;
                //CA410_Measurement.AutoConnect();
                //CA410_Measurement.DefaultSetting();
                //CA410_Measurement.chnum = int.Parse(CA410_Channel.Text);
                //CA410_Measurement.Measurement();
                //MessageBox.Show("Connect CA410 Success!", "连接信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Connect CA410 Fail!", "连接信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        /// <summary>
        /// 关闭窗口触发事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("是否退出?", "消息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            if (dr == DialogResult.OK)   //如果单击“是”按钮
            {
                e.Cancel = false;                 //关闭窗体

                para[11] = 0;
                para[12] = 0;
                para[13] = 0;
                para[14] = 0;
                para[15] = 0;
                para[16] = 0;
                para[17] = 0;
                para[18] = 0;
                para[19] = 0;
                para[20] = 0;
                Parawrite(para, para.Length);
            }
            else if (dr == DialogResult.Cancel)
            {
                e.Cancel = true;                  //不执行操作
                //MessageBox.Show("已取消关闭窗口", "消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //------------------------------------------USB-----------------------------------------------------------------------------------------------------------------------------------
        public delegate void SendPara(byte[] data, int length);
        public SendPara SendataHandle = null;

        /// <summary>
        /// 写参数到C-B
        /// </summary>
        /// <param name="wr_addr"></param>
        /// <param name="wr_length"></param>
        public void Parawrite(byte[] para, int length)
        {
            if (SendataHandle != null)
            {
                SendataHandle(para, para.Length);
            }
        }

        /// <summary>
        /// USB更新
        /// </summary>
        public void UsbDeviceHandle()
        {
            usb.usbDeviceState += UsbStateUpdate;
            wct_rgb.SendataHandle += sendpara;
            SendataHandle += sendpara;
        }
        /// <summary>
        /// USB连接状态更新
        /// </summary>
        public void UsbStateUpdate()
        {
            Connect_Status_Label.Text = usb.connect;
            toolStrip1.BackColor = System.Drawing.Color.FromArgb(usb.Color_Red, usb.Color_Green, usb.Color_Blue);
        }
        /// <summary>
        /// 发送参数
        /// </summary>
        /// <param name="data"></param>
        /// <param name="length"></param>
        private void sendpara(byte[] data, int length)
        {
            result = usb.SendTransfers(data, length);
            if (result)
            {
                //MessageBox.Show("参数发送成功");
                Console.WriteLine("参数发送成功");
            }
            else
            {
                MessageBox.Show("参数发送失败");
                //Console.WriteLine("参数发送失败");
            }
        }

        //---------------------------------------Digital Gamma---------------------------------------------------------------------------------------------------------------------------
        public void ParaPassing()
        {
            digitalgamma.RGB = rBtn_RGB.Checked;
            digitalgamma.m_x = Convert.ToDouble(nUD_mx.Value);
            digitalgamma.m_y = Convert.ToDouble(nUD_my.Value);
            digitalgamma.tarlum = Convert.ToDouble(nUD_tarlum.Value);
            digitalgamma.gamma = Convert.ToDouble(nUD_gamma.Value);

            if (rBtn_data10.Checked)
            {
                digitalgamma.datawidth = 1024;
            }
            else
            {
                digitalgamma.datawidth = 4096;
            }

            if (rBtn_adr8.Checked)
            {
                digitalgamma.addresswidth = 256;
            }
            else if (rBtn_adr10.Checked)
            {
                digitalgamma.addresswidth = 1024;
            }
            else if (rBtn_adr11.Checked)
            {
                digitalgamma.addresswidth = 2048;
            }
            else
            {
                digitalgamma.addresswidth = 0;
            }
        }
        private void btn_loadxlsx_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "(*.xlsx)|*.xlsx|" + "(*.xls)|*.xls|" + "(*.*)|*.*";
                if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
                {
                    digitalgamma.dt = ReadExcelToTable(openFileDialog1.FileName);
                    ParaPassing();
                    digitalgamma.DataCache(digitalgamma.dt);
                    digitalgamma.MaxLumLimit();
                    nUD_tarlum.Maximum = digitalgamma.maxlum_limit;
                    nUD_tarlum.Value = digitalgamma.maxlum_limit;
                    //nUD_α.Text = digitalgamma.α.ToString("f4");
                    //nUD_β.Text = digitalgamma.β.ToString("f4");
                    //nUD_γ.Text = digitalgamma.γ.ToString("f4");
                    ShowMessage();
                    MessageBox.Show("基于R、G、B最高亮度将目标亮度最大值限制为<" + nUD_tarlum.Maximum + "nits>", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    MessageBox.Show("读取失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("格式错误", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
         
        public static DataTable ReadExcelToTable(string path)//excel存放的路径
        {
            try
            {
                //连接字符串
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });         //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串                                              //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set);
                    return set.Tables[0];
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public void ShowMessage()
        {
            ParaPassing();
            lab_rx.Text = digitalgamma.dt.Rows[257][4].ToString();
            lab_ry.Text = digitalgamma.dt.Rows[257][5].ToString();
            lab_rmaxlum.Text = digitalgamma.dt.Rows[257][6].ToString();
            lab_gx.Text = digitalgamma.dt.Rows[513][4].ToString();
            lab_gy.Text = digitalgamma.dt.Rows[513][5].ToString();
            lab_gmaxlum.Text = digitalgamma.dt.Rows[513][6].ToString();
            lab_bx.Text = digitalgamma.dt.Rows[769][4].ToString();
            lab_by.Text = digitalgamma.dt.Rows[769][5].ToString();
            lab_bmaxlum.Text = digitalgamma.dt.Rows[769][6].ToString();
            ShowMessage2();

            if (rBtn_RGB.Checked)
            {
                lab_wx.Text = "-";
                lab_wy.Text = "-";
                lab_wmaxlum.Text = "-";
            }
            else
            {
                lab_wx.Text = digitalgamma.dt.Rows[1025][4].ToString();
                lab_wy.Text = digitalgamma.dt.Rows[1025][5].ToString();
                lab_wmaxlum.Text = digitalgamma.dt.Rows[1025][6].ToString();
            }
        }

        public void ShowMessage2()
        {
            ParaPassing();
            digitalgamma.LumFinalCal();
            lab_rlumfinal.Text = digitalgamma.r_lumfinal.ToString("f4");
            lab_glumfinal.Text = digitalgamma.g_lumfinal.ToString("f4");
            lab_blumfinal.Text = digitalgamma.b_lumfinal.ToString("f4");
            if (rBtn_RGB.Checked)
            {
                lab_wlumfinal.Text = "-";
            }
            else
            {
                lab_wlumfinal.Text = digitalgamma.w_lumfinal.ToString("f4");
            }
        }
        private void btn_saveDiGaLUT_Click(object sender, EventArgs e)
        {
            try
            {
                ParaPassing();
                digitalgamma.LUTGen();
                System.Windows.Forms.SaveFileDialog objSave = new System.Windows.Forms.SaveFileDialog();
                objSave.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
                objSave.FileName = "Digital_Gamma_LUT_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".bin";
                if (objSave.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = new FileStream(objSave.FileName, FileMode.OpenOrCreate);
                    BinaryWriter binWriter = new BinaryWriter(fs);
                    binWriter.Write(digitalgamma.flash, 0, digitalgamma.flash.Length);
                    binWriter.Close();
                    fs.Close();
                }
                else
                {
                    MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("格式错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void nUD_mx_TextChanged(object sender, EventArgs e)
        {
            ParaPassing();
            digitalgamma.MaxLumLimit();
            nUD_tarlum.Maximum = digitalgamma.maxlum_limit;
            nUD_tarlum.Value = digitalgamma.maxlum_limit;
            nUD_α.Text = digitalgamma.α.ToString("f4");
            nUD_β.Text = digitalgamma.β.ToString("f4");
            nUD_γ.Text = digitalgamma.γ.ToString("f4");
            ShowMessage2();
        }

        private void nUD_my_TextChanged(object sender, EventArgs e)
        {
            ParaPassing();
            digitalgamma.MaxLumLimit();
            nUD_tarlum.Maximum = digitalgamma.maxlum_limit;
            nUD_tarlum.Value = digitalgamma.maxlum_limit;
            nUD_α.Text = digitalgamma.α.ToString("f4");
            nUD_β.Text = digitalgamma.β.ToString("f4");
            nUD_γ.Text = digitalgamma.γ.ToString("f4");
            ShowMessage2();
        }

        private void nUD_tarlum_TextChanged(object sender, EventArgs e)
        {
            ShowMessage2();
        }

        private void rBtn_adr8_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn_adr8.Checked)
            {
                rBtn_data10.Checked = true;
                rBtn_data12.Checked = false;
            }
            else if (rBtn_adr10.Checked)
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
            else
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
        }

        private void rBtn_adr10_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn_adr8.Checked)
            {
                rBtn_data10.Checked = true;
                rBtn_data12.Checked = false;
            }
            else if (rBtn_adr10.Checked)
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
            else
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
        }

        private void rBtn_adr11_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn_adr8.Checked)
            {
                rBtn_data10.Checked = true;
                rBtn_data12.Checked = false;
            }
            else if (rBtn_adr10.Checked)
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
            else
            {
                rBtn_data10.Checked = false;
                rBtn_data12.Checked = true;
            }
        }

        private void RBtn_adr11_CheckedChanged(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        //----------------------------------------------CLA---------------------------------------------
        private void btn_load_claxlsx_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "(*.xlsx)|*.xlsx|" + "(*.xls)|*.xls|" + "(*.*)|*.*";
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                DataTable dt_CLA = ReadExcelToTable(openFileDialog1.FileName);
                dGV_CLA.DataSource = dt_CLA;
                cla.dt_CLA = dt_CLA;
            }
            else
            {
                MessageBox.Show("读取失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_loadDIGAlut_Click(object sender, EventArgs e)
        {
            byte[] readflash = new byte[16390];
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.OpenOrCreate);
                BinaryReader binReader = new BinaryReader(fs);
                binReader.Read(readflash, 0, 16390);
                cla.LUTRead(readflash);
                binReader.Close();
                fs.Close();
            }
            else
            {
                MessageBox.Show("加载失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_saveCLAlut_Click(object sender, EventArgs e)
        {
            try
            {
                cla.Imax = Convert.ToDouble(nUD_Imax.Text);
                cla.LUTWrite();
                System.Windows.Forms.SaveFileDialog objSave = new System.Windows.Forms.SaveFileDialog();
                objSave.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
                objSave.FileName = "CLA_LUT_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".bin";
                if (objSave.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = new FileStream(objSave.FileName, FileMode.OpenOrCreate);
                    BinaryWriter binWriter = new BinaryWriter(fs);
                    binWriter.Write(cla.writeflash, 0, cla.writeflash.Length);
                    binWriter.Close();
                    fs.Close();
                }
                else
                {
                    MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("格式错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        //--------------------------------------------WCT-RGBW------------------------------------------
        public void paratras()
        {
            para[0] = (byte)(Convert.ToDouble(nUD_lum.Value) % 256);
            para[1] = (byte)(Convert.ToDouble(nUD_lum.Value) / 256);
            para[2] = (byte)(Convert.ToDouble(nUD_α.Value) * 65536 % 256);
            para[3] = (byte)(Convert.ToDouble(nUD_α.Value) * 65536 / 256);
            para[4] = (byte)(Convert.ToDouble(nUD_β.Value) * 65536 % 256);
            para[5] = (byte)(Convert.ToDouble(nUD_β.Value) * 65536 / 256);
            para[6] = (byte)(Convert.ToDouble(nUD_γ.Value) * 65536 % 256);
            para[7] = (byte)(Convert.ToDouble(nUD_γ.Value) * 65536 / 256);
        }

        private void parawrite(int st_addr, int length)
        {
            paratras();

            wbuff[0] = 0x00;
            wbuff[1] = DeviceAdder;
            wbuff[2] = (byte)length;
            wbuff[3] = (byte)(st_addr >> 8);
            wbuff[4] = (byte)(st_addr);
            for (int i = 1; i <= length; i++)
            {
                wbuff[i + 4] = para[st_addr + i - 1];
            }

            hid_report = new report(0x00, wbuff);

            //检测设备是否已经找到或打开
            if (hidReturn == Hid.HID_RETURN.SUCCESS || hidReturn == Hid.HID_RETURN.DEVICE_OPENED)
            {
                //发送数据
                switch (hidUsb.Write(hid_report))
                {
                    case Hid.HID_RETURN.WRITE_FAILD:
                        MessageBox.Show("Write Failed ！");
                        break;
                        //case Hid.HID_RETURN.SUCCESS:
                        //    MessageBox.Show("Write Successed !");
                        //    break;
                }
            }
            else
            {
                MessageBox.Show("Please Find The Device ！");
            }
        }
        private void btn_i2cwrite_Click(object sender, EventArgs e)
        {
            int st_addr = 0;
            int length = 8;
            parawrite(st_addr, length);

            //bool result = usb.SendTransfers(para, para.Length);
            //if (result)
            //{
            //    MessageBox.Show("参数发送成功");
            //}
            //else
            //{
            //    MessageBox.Show("参数发送失败");
            //}
        }

        private void btn_connect_Click(object sender, EventArgs e)
        {
            hidReturn = hidUsb.OpenDevice(vendorId, productorId, "");

            switch (hidReturn)
            {
                case Hid.HID_RETURN.SUCCESS:
                    MessageBox.Show("Open Success!");
                    break;
                case Hid.HID_RETURN.DEVICE_OPENED:
                    MessageBox.Show("Device Opened!");
                    break;
                case Hid.HID_RETURN.DEVICE_NOT_FIND:
                    MessageBox.Show("Device Not Find!");
                    break;
                case Hid.HID_RETURN.NO_DEVICE_CONECTED:
                    MessageBox.Show("No Device Conected!");
                    break;
                case Hid.HID_RETURN.READ_FAILD:
                    MessageBox.Show("Read Faild!");
                    break;
                case Hid.HID_RETURN.WRITE_FAILD:
                    MessageBox.Show("Write Faild!");
                    break;
            }
        }

        private void btn_loadCLAlut_Click(object sender, EventArgs e)
        {
            byte[] readflash = new byte[32774];
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.OpenOrCreate);
                BinaryReader binReader = new BinaryReader(fs);
                binReader.Read(readflash, 0, 32774);
                dGV_WCT.DataSource = null; //每次打开清空内容
                wct.LUTRead(readflash);
                dGV_WCT.DataSource = wct.dt_rgbcomp;
                binReader.Close();
                fs.Close();
            }
            else
            {
                MessageBox.Show("加载失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_lutwrite_Click(object sender, EventArgs e)
        {
            wct.dt_rgbcomp.Rows[hSB_lum.Value][0] = "*";
            wct.dt_rgbcomp.Rows[hSB_lum.Value][1] = hSB_lum.Value;
            wct.dt_rgbcomp.Rows[hSB_lum.Value][2] = (double)hSB_α.Value / 10000;
            wct.dt_rgbcomp.Rows[hSB_lum.Value][3] = (double)hSB_β.Value / 10000;
            wct.dt_rgbcomp.Rows[hSB_lum.Value][4] = (double)hSB_γ.Value / 10000;
            wct.lum = hSB_lum.Value;
            wct.TableChange();
            dGV_WCT.DataSource = wct.dt_rgbcomp;
        }

        private void btn_saveWCTlut_Click(object sender, EventArgs e)
        {
            try
            {
                wct.LUTWrite();
                System.Windows.Forms.SaveFileDialog objSave = new System.Windows.Forms.SaveFileDialog();
                objSave.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
                objSave.FileName = "WCT_RGB_LUT_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".bin";
                if (objSave.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = new FileStream(objSave.FileName, FileMode.OpenOrCreate);
                    BinaryWriter binWriter = new BinaryWriter(fs);
                    binWriter.Write(wct.writeflash, 0, wct.writeflash.Length);
                    binWriter.Close();
                    fs.Close();
                }
                else
                {
                    MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("格式错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void hSB_lum_Scroll(object sender, ScrollEventArgs e)
        {
            nUD_lum.Value = hSB_lum.Value;
        }
        private void hSB_α_Scroll(object sender, ScrollEventArgs e)
        {
            nUD_α.Text = (((double)hSB_α.Value) / 10000).ToString("f4");
        }
        private void hSB_β_Scroll(object sender, ScrollEventArgs e)
        {
            nUD_β.Text = (((double)hSB_β.Value) / 10000).ToString("f4");
        }
        private void hSB_γ_Scroll(object sender, ScrollEventArgs e)
        {
            nUD_γ.Text = (((double)hSB_γ.Value) / 10000).ToString("f4");
        }

        private void nUD_lum_ValueChanged(object sender, EventArgs e)
        {
            hSB_lum.Value = (int)nUD_lum.Value;
            //int st_addr = 0;
            //int length = 2;
            //parawrite(st_addr, length);
        }

        private void nUD_α_ValueChanged(object sender, EventArgs e)
        {
            hSB_α.Value = (int)(nUD_α.Value * 10000);
            //int st_addr = 2;
            //int length = 2;
            //parawrite(st_addr, length);
        }

        private void nUD_β_ValueChanged(object sender, EventArgs e)
        {
            hSB_β.Value = (int)(nUD_β.Value * 10000);
            //int st_addr = 4;
            //int length = 2;
            //parawrite(st_addr, length);
        }

        private void nUD_γ_ValueChanged(object sender, EventArgs e)
        {
            hSB_γ.Value = (int)(nUD_γ.Value * 10000);
            //int st_addr = 6;
            //int length = 2;
            //parawrite(st_addr, length);
        }
        //--------------------------------------------WCT-RGB-------------------------------------------
        /// <summary>
        /// 连接到410
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Connect_310_410_Click(object sender, EventArgs e)
        {
            //try
            //{
                para[11] = 1;               //黑画面以便CA410校零
                para[12] = 0;
                para[13] = 0;
                para[14] = 0;
                para[15] = 0;
                para[16] = 0;
                para[17] = 0;
                para[18] = 0;
                para[19] = 0;
                para[20] = 0;
                Parawrite(para, para.Length);

                CA410_Measurement.AutoConnect();
                CA410_Measurement.DefaultSetting();
                CA410_Measurement.chnum = int.Parse(CA410_Channel.Text);
                CA410_Measurement.Measurement();
                //MessageBox.Show("Connect CA410 Success!", "连接信息", MessageBoxButtons.OK, MessageBoxIcon.Information);

                para[11] = 0;               //CA410校零完成
                para[12] = 0;
                para[13] = 0;
                para[14] = 0;
                para[15] = 0;
                para[16] = 0;
                para[17] = 0;
                para[18] = 0;
                para[19] = 0;
                para[20] = 0;
                Parawrite(para, para.Length);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Connect CA410 Fail!", "连接信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        /// <summary>
        /// 获取当前亮度色坐标信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Get_Current_Lum_Click(object sender, EventArgs e)
        {
            try
            {
                CA410_Measurement.Measurement();
                Console.WriteLine("亮度:" + CA410_Measurement.Lv + "色坐标X:" + CA410_Measurement.sx + "色坐标Y:" + CA410_Measurement.sy);
                //dGV_Cur_Lum.DataSource = CA410_Measurement.Lv;
                Dt_Cur_Lum.Rows[0][0] = CA410_Measurement.Lv.ToString("F4");
                Dt_Cur_X.Rows[0][0] = CA410_Measurement.sx.ToString("F4");
                Dt_Cur_Y.Rows[0][0] = CA410_Measurement.sy.ToString("F4");

                dGV_Cur_Lum.DataSource = Dt_Cur_Lum;
                dGV_Cur_X.DataSource = Dt_Cur_X;
                dGV_Cur_Y.DataSource = Dt_Cur_Y;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Confirm CA410 Connection!", "连接信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 移除窗体对象
        /// </summary>
        /// <param name="yyyy-MM-dd HH-mm-ss"></param>
        public void RemoveForm(Form f)
        {
            if (allforms.Contains(f))
            {
                allforms.Remove(f);
            }
        }

        /// <summary>
        /// 选择410通道
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CA410_Channel_SelectedIndexChanged(object sender, EventArgs e)
        {
            CA410_Measurement.chnum = int.Parse(CA410_Channel.Text);

        }


        /// <summary>
        /// 创建RGB_Pattern测量亮度色坐标
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RGB_Pattern_Btn_Click(object sender, EventArgs e)
        {
            try
            {
                //if (result == false)
                //{
                //    btn_RGB_Pattern_Btn.Enabled = true;
                //    MessageBox.Show("参数发送失败");
                //    return;
                //}
                //else
                //{
                //    btn_RGB_Pattern_Btn.Enabled = false;
                //    MessageBox.Show("参数发送成功");
                //}
                //btn_RGB_Pattern_Btn.Enabled = false;

                dGV_Trans_Num.DataSource = wct_rgb.G2V_Trans_Num_Tab;
                dGV_Cur_Lum.DataSource = wct_rgb.G2V_Current_Lum;
                dGV_Cur_X.DataSource = wct_rgb.G2V_Current_X;
                dGV_Cur_Y.DataSource = wct_rgb.G2V_Current_Y;
                dGV_Max_Lum_Limit.DataSource = wct_rgb.G2V_Max_Lum_Limit;
                dGV_Disp_G2V_Table.DataSource = wct_rgb.G2V_410MeasData;

                ////数据小于目标值则显示为红色
                //if (Convert.ToDouble(Convert.ToDouble(wct_rgb.G2V_Max_Lum_Limit.Rows[0][0])) < Convert.ToDouble(Targ_Lum.Text))
                //{
                //    dGV_Max_Lum_Limit.ForeColor = System.Drawing.Color.FromArgb(usb.Color_Red, usb.Color_Green, usb.Color_Blue);
                //}

                Tar_Lum = Targ_Lum_Text.Text;
                Tar_x = Targ_X_Text.Text;
                Tar_y = Targ_Y_Text.Text;

                Th_Tran_RGBPattern_To_Panel = new Thread(wct_rgb.TransGrayPatternToC_B);
                Th_Tran_RGBPattern_To_Panel.IsBackground = true;
                Th_Tran_RGBPattern_To_Panel.Start();

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message, "创造Pattern", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox_TestGraySel_SelectedIndexChanged(object sender, EventArgs e)
        {
            TestGray = comboBox_TestGraySel.Text;
        }

        private void btn_TestGrayTrans_Click(object sender, EventArgs e)
        {
            TestGray = comboBox_TestGraySel.Text;
            byte[] para = new byte[512];
            switch (TestGray)
            {
                case "R":
                    para[11] = 1;
                    para[12] = 0;
                    para[13] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text));
                    para[14] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text) / 256);
                    para[15] = 0;
                    para[16] = 0;
                    para[17] = 0;
                    para[18] = 0;
                    para[19] = 0;
                    para[20] = 0;
                    Console.WriteLine("R");
                    break;
                case "G":
                    para[11] = 1;
                    para[12] = 0;
                    para[13] = 0;
                    para[14] = 0;
                    para[15] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text));
                    para[16] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text) / 256);
                    para[17] = 0;
                    para[18] = 0;
                    para[19] = 0;
                    para[20] = 0;
                    Console.WriteLine("G");
                    break;
                case "B":
                    para[11] = 1;
                    para[12] = 0;
                    para[13] = 0;
                    para[14] = 0;
                    para[15] = 0;
                    para[16] = 0;
                    para[17] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text));
                    para[18] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text) / 256);
                    para[19] = 0;
                    para[20] = 0;
                    Console.WriteLine("B");
                    break;
                default:
                    para[11] = 1;
                    para[12] = 0;
                    para[13] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text));
                    para[14] = (byte)(Convert.ToInt16(nUpDown_GrayTransTest.Text) / 256);
                    para[15] = 0;
                    para[16] = 0;
                    para[17] = 0;
                    para[18] = 0;
                    para[19] = 0;
                    para[20] = 0;
                    Console.WriteLine("default");
                    break;

            }
            Parawrite(para, para.Length);
        }




        private void btn_Stop_Trans_Pattern_Click(object sender, EventArgs e)
        {
            try
            {
                Th_Tran_RGBPattern_To_Panel.Abort();
                btn_RGB_Pattern_Btn.Enabled = true;
                MessageBox.Show("停止发送Pattern！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("当前无线程！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Save_G2V_LUT_Click(object sender, EventArgs e)
        {
            G2V_Adr_Data_Width_Setting();
            dGV_Disp_G2V_Table.DataSource = wct_rgb.wct_dt;
            wct_rgb.LUTWrite();

            System.Windows.Forms.SaveFileDialog objSave = new System.Windows.Forms.SaveFileDialog();
            objSave.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            objSave.FileName = "WCT_RGB_LUT_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".bin";
            if (objSave.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(objSave.FileName, FileMode.OpenOrCreate);
                BinaryWriter binWriter = new BinaryWriter(fs);
                binWriter.Write(wct_rgb.writeflash, 0, wct_rgb.writeflash.Length);
                binWriter.Close();
                fs.Close();
                MessageBox.Show("保存成功！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Load_DGma_Lut_Click(object sender, EventArgs e)
        {
            byte[] readflash = new byte[16392];          //8 Page Digamma data 8*2048 = 16384
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.OpenOrCreate);
                BinaryReader binReader = new BinaryReader(fs);
                binReader.Read(readflash, 0, 16392);
                dGV_Disp_G2V_Table.DataSource = null; //每次打开清空内容
                wct_rgb.LUTRead(readflash);
                dGV_Disp_G2V_Table.DataSource = wct_rgb.wct_dt;
                binReader.Close();
                fs.Close();
            }
            else
            {
                MessageBox.Show("加载失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Compbox_Gray_Pattern_Num_SelectedIndexChanged(object sender, EventArgs e)
        {
            wct_rgb.G2V_Gray_Pattern_Num = Convert.ToInt16(Compbox_Gray_Pattern_Num.Text);
        }

        private void Adjust_G2V_LUT_Click(object sender, EventArgs e)
        {
            dGV_Correct_Num.DataSource = wct_rgb.G2V_Correct_Num_Tab; //当前校正灰阶

            dGV_Cur_Lum.DataSource = wct_rgb.G2V_Current_Lum;         //当前亮度，色坐标
            dGV_Cur_X.DataSource = wct_rgb.G2V_Current_X;
            dGV_Cur_Y.DataSource = wct_rgb.G2V_Current_Y;

            dGV_Disp_G2V_Table.DataSource = wct_rgb.wct_dt;

            G2V_Adr_Data_Width_Setting();                            //LUT地址数据位宽设置

            //if (result == false)
            //{
            //    btn_Adjust_G2V_LUT.Enabled = true;
            //    MessageBox.Show("参数发送失败");
            //    return;
            //}
            //else
            //{
            //    btn_Adjust_G2V_LUT.Enabled = false;
            //    MessageBox.Show("参数发送成功");
            //}
            //btn_Adjust_G2V_LUT.Enabled = false;
            Th_G2V_Correct = new Thread(wct_rgb.G2V_LUT_Correction);
            Th_G2V_Correct.IsBackground = true;
            Th_G2V_Correct.Start();
        }

        private void btn_G2V_Cor_Abort_Click(object sender, EventArgs e)
        {
            try
            {
                Th_G2V_Correct.Abort();
                btn_Adjust_G2V_LUT.Enabled = true;
                MessageBox.Show("G2V校正停止", "信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("当前无线程！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void nUpDown_Std_Lum_ValueChanged(object sender, EventArgs e)
        {
            wct_rgb.Std_Lum = (double)nUpDown_Std_Lum.Value;
            //Console.WriteLine("亮度标准改变为：" + wct_rgb.Std_Lum);
        }

        private void nUpDown_Std_Y_ValueChanged(object sender, EventArgs e)
        {
            wct_rgb.Std_Y = (double)nUpDown_Std_Y.Value;
            //Console.WriteLine("X标准改变为：" + wct_rgb.Std_Y);
        }

        private void nUpDown_Std_X_ValueChanged(object sender, EventArgs e)
        {
            wct_rgb.Std_X = (double)nUpDown_Std_X.Value;
            //Console.WriteLine("Y标准改变为：" + wct_rgb.Std_X);
        }
        private void G2V_Adr_Data_Width_Setting()
        {
            if (rbtn_G2V_adr8.Checked)
            {
                wct_rgb.AddressWidth = 256;
                Console.WriteLine("地址位宽：" + 256);
            }
            else if (rbtn_G2V_adr10.Checked)
            {
                Console.WriteLine("地址位宽：" + 1024);
                wct_rgb.AddressWidth = 1024;
            }
            else
            {
                wct_rgb.AddressWidth = 2048;
                Console.WriteLine("地址位宽：" + 2048);
            }

            if (rbtn_G2V_data10.Checked)
            {
                wct_rgb.DataWidth = 1024;
                Console.WriteLine("数据位宽：" + 1024);
            }
            else if (rbtn_G2V_data11.Checked)
            {
                wct_rgb.DataWidth = 2048;
                Console.WriteLine("数据位宽：" + 2048);
            }
            else
            {
                wct_rgb.DataWidth = 4096;
                Console.WriteLine("数据位宽：" + 4096);
            }
        }

        /// <summary>
        /// 加载多个bin档
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Load_Multiple_Bin_Click(object sender, EventArgs e)
        {
            Read_Bin_File_Num = Convert.ToInt16(combBox_Bin_File_Num.Text);
            OpenFileDialog Open_Dlg = new OpenFileDialog();
            Open_Dlg.Multiselect = true;
            Open_Dlg.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            if (Open_Dlg.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < Read_Bin_File_Num; i++)
                {
                    FileStream fs = new FileStream(Open_Dlg.FileNames[i], FileMode.OpenOrCreate);    //加载多个文件，按文件名字符串中顺序读取文件
                    BinaryReader binReader = new BinaryReader(fs);
                    binReader.Read(Read_G2V_LUT, 16384 * i, 16384);
                    binReader.Close();
                    fs.Close();
                }

                MessageBox.Show("加载成功！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("加载失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// 改变加载个数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void combBox_Bin_File_Num_SelectedIndexChanged(object sender, EventArgs e)
        {
            Read_Bin_File_Num = Convert.ToInt16(combBox_Bin_File_Num.Text);
        }

        /// <summary>
        /// 合并多个bin档为一个
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Save_AS_One_Bin_Click(object sender, EventArgs e)
        {
            Read_Bin_File_Num = Convert.ToInt16(combBox_Bin_File_Num.Text);
            SaveFileDialog Save_Dlg = new System.Windows.Forms.SaveFileDialog();
            Save_Dlg.Filter = "(*.bin)|*.bin|" + "(*.*)|*.*";
            Save_Dlg.FileName = "G2V_LUT合并_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".bin";

            if (Save_Dlg.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(Save_Dlg.FileName, FileMode.OpenOrCreate);
                BinaryWriter binWriter = new BinaryWriter(fs);
                binWriter.Write(Read_G2V_LUT, 0, 16384 * Read_Bin_File_Num);
                binWriter.Close();
                fs.Close();
                MessageBox.Show("保存成功！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox_MeasInterval_SelectedIndexChanged(object sender, EventArgs e)
        {
            wct_rgb.Delay = Convert.ToInt16(comboBox_MeasInterval.Text);
        }

        /// <summary>
        /// Data写到Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Path"></param>
        public void Dt2Excel(DataTable dt, string Path)
        {
            string StrCon = string.Empty;
            FileInfo file = new FileInfo(Path);
            string extension = file.Extension;
            string tablename = "CA410MeasData";
            switch (extension)
            {
                case ".xls":
                    StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0;'";
                    //strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0;'";
                    //strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=2;'";
                    break;
                case ".xlsx":
                    //strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;";
                    //strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2;'";    //出现错误了
                    StrCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties= 'Excel 12.0;HDR=Yes;IMEX=0;'"; 
                    break;
                default:
                    StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0;'";
                    break;
            }
            try
            {
                using (System.Data.OleDb.OleDbConnection Con = new System.Data.OleDb.OleDbConnection(StrCon))
                {

                    StringBuilder StrSQL_CreateTable = new StringBuilder();
                    StringBuilder StrSQL_Add_Column = new StringBuilder();
                    StringBuilder StrSQL_Add_Row = new StringBuilder();
                    System.Data.OleDb.OleDbCommand Cmd;

                    Con.Close();
                    Con.Open();
                    

                    //创建表格
                    StrSQL_CreateTable.Append("CREATE TABLE ").Append("[" + tablename + "] ");
                    //添加列
                    StrSQL_CreateTable.Append("(");
                    StrSQL_CreateTable.Append("[Gray] string,");
                    StrSQL_CreateTable.Append("[Gray_Num] int,");
                    StrSQL_CreateTable.Append("[Lum] double,");
                    StrSQL_CreateTable.Append("[X] double,");
                    StrSQL_CreateTable.Append("[Y] double");
                    StrSQL_CreateTable.Append(");");

                    Cmd = new System.Data.OleDb.OleDbCommand(StrSQL_CreateTable.ToString(), Con);
                    Cmd.ExecuteNonQuery();

                //添加行
                for (int i = 0; i < dt.Rows.Count; i++)                    //INSERT INTO 表名称 VALUES (Gray,Gray_Num),......(Gray,Gray_Num);
                {
                    StrSQL_CreateTable.Clear();
                    StrSQL_Add_Row.Append("(");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        StrSQL_Add_Row.Append("'" + dt.Rows[i][j] + "'");
                        if (j < dt.Columns.Count - 1)
                        {
                            StrSQL_Add_Row.Append(",");
                        }
                        else
                        {
                            //if (i == dt.Rows.Count - 1)
                            //{
                            //StrSQL_Add_Row.Append(");");
                            //}
                            //else
                            //{
                            StrSQL_Add_Row.Append(");");
                            //}
                        }
                    }
                    //Microsoft Access 数据写到Excel只能逐行写入，故只能在每行数据完成时写入并清除以备下一行使用，后续有机会改进
                    Cmd.CommandText = StrSQL_CreateTable.Append("INSERT INTO [" + tablename + "] VALUES ").Append(StrSQL_Add_Row).ToString();
                    Cmd.ExecuteNonQuery();
                    StrSQL_Add_Row.Clear();  //清除以备下一行使用
                }
                Con.Close();
                    Con.Dispose();
                    MessageBox.Show("保存成功", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        /// <summary>
        /// 保存量测到的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Save_410MeasData_Click(object sender, EventArgs e)
        {
            SaveFileDialog Save_Dlg = new SaveFileDialog();
            Save_Dlg.Filter = "(*.xls)|*.xls|" + "(*.xlsx)|*.xlsx|" + "(*.*)|*.*";
            Save_Dlg.FileName = "CA410MeasData_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".xls";
            //Save_Dlg.FileName = "CA410MeasData.xls";
            
            if (Save_Dlg.ShowDialog() == DialogResult.OK)
            {

                if (File.Exists(Save_Dlg.FileName))
                {
                    File.Delete(Save_Dlg.FileName);
                }
                else
                {
                    //File.Create(Save_Dlg.FileName);   //这里可能导致数据库引擎打不开
                }
                Dt2Excel(wct_rgb.G2V_410MeasData, Save_Dlg.FileName);
                
            }
            else
            {
                MessageBox.Show("保存失败！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
    //---------------------------------------Para-Setting----------------------------------------------------
}


