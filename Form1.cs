using AForge.Controls;
using AForge.Imaging.Filters;
using AForge.Video.DirectShow;
using Baidu.Aip.Ocr;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using Rectangle = System.Drawing.Rectangle;
using System.Web;
using System.Collections;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Speech.Recognition;
using System.Timers;

namespace WindowsFormsApp1
{

   
        public partial class Form1 : Form
    {

        /// <summary>
        /// 百度OCR，个人申请的免费在线OCR
        /// </summary>
        string API_KEY = "Ahkop6uNoQ7f3YYgocvOw4t3";
        string SECRET_KEY = "INnI66eKUHNdivdh08jGhI1NKzwdfeQm";
        Ocr client;
        private FilterInfoCollection videoDevices;

        public Form1()
        {
            InitializeComponent();
        }

        System.Timers.Timer t = new System.Timers.Timer(5000); //设置时间间隔为5秒
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“db_mytestDataSet.个人信息”中。您可以根据需要移动或移除它。
            client = new Baidu.Aip.Ocr.Ocr(API_KEY, SECRET_KEY);
            client.Timeout = 2000;  // 修改超时时间

            //定时器设置
            t.Elapsed += new System.Timers.ElapsedEventHandler(Timer_TimesUp);
            t.AutoReset = true; //每到指定时间Elapsed事件是触发一次（false），还是一直触发（true）
          
            try
            {
                // 枚举所有视频输入设备
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                if (videoDevices.Count == 0)
                    throw new ApplicationException();
               foreach (FilterInfo device in videoDevices)
                {
                    comboBox1.Items.Add(device.Name);
                }
                comboBox1.SelectedIndex = 0;

            }
            catch (ApplicationException)
            {
                comboBox1.Items.Add("未发现摄像头");
                comboBox1.SelectedIndex = 0;
                videoDevices = null;
            }

        }
      
       
        public string GeneralBasicDemo(string filename)
        {
            var image = File.ReadAllBytes(filename);
            recglized(image);
            return null;
          
        }
        private void recglized(byte[] imag)
        {
            var options = new Dictionary<string, object>{
                {"language_type", "CHN_ENG"},
                {"detect_direction", "true"},
                {"detect_language", "true"},
                {"probability", "true"}
                                                        };

            var result = client.GeneralBasic(imag, options);
            string str = "正在识别....";
            var txts = (from obj in (JArray)result.Root["words_result"]
                        select (string)obj["words"]);
            List<xinxi> list = new List<xinxi>();
            xinxi xx = new xinxi();

            bool rowend = false;
            foreach (var r in txts)
            {
                if (rowend == true) xx = new xinxi();

                if (Regex.IsMatch(r, @"^\d{3}$"))
                {
                    xx.n0 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"^\w复归$"))
                {
                    xx.n1 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"确认$"))
                {
                    xx.n2 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"\d{4,}"))
                {
                    xx.n3 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"\w{8,}"))
                {
                    xx.n4 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"异常"))
                {
                    xx.n7 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"\w(.*)日(.*)$"))
                {
                    xx.n8 = r;
                    rowend = false;
                    continue;
                }
                if (Regex.IsMatch(r, @"^遥信变位$"))
                {
                    xx.n9 = r;
                    rowend = true;
                    list.Add(xx);

                }
            }      
            richTextBox1.AppendText(result.ToString());
            dataGridView1.AutoGenerateColumns = false;
          if(list.Count!=0)
            {
              dataGridView1.DataSource = list;
            }
           

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                GeneralBasicDemo(textBox1.Text);
            }
        }

        [Obsolete]
        private void button2_Click(object sender, EventArgs e)
        {
            //获取摄像头
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            //实例化摄像头
            VideoCaptureDevice videoDevice = new VideoCaptureDevice(videoDevices[comboBox1.SelectedIndex].MonikerString);
            //将摄像头视频播放在控件中
            videoSourcePlayer1.VideoSource = videoDevice;
            //开启摄像头
            videoSourcePlayer1.Start();
         
        }
      

        private void button3_Click(object sender, EventArgs e)
        {
            if (videoSourcePlayer1 != null && videoSourcePlayer1.IsRunning)
            {
                videoSourcePlayer1.SignalToStop();
                videoSourcePlayer1.WaitForStop();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void videoSourcePlayer1_NewFrame(object sender, ref Bitmap image)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //拍照
           
            Bitmap img = videoSourcePlayer1.GetCurrentVideoFrame();
         
            pictureBox1.Image = (Image)img.Clone();
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            recglized(BitmapToGrayByte(img));
            pictureBox1.Image.Dispose();
            
            //保存文件
           /* string path = System.Windows.Forms.Application.StartupPath + "\\image";//根目录下的image文件夹
            if (Directory.Exists(path) == false)
            {//判断目录是否存在
                Directory.CreateDirectory(path);
            }
            string fileName = "img" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".png";//给照片文件命名
            img.Save(path + "\\" + fileName);//保存照片文件，其中image是摄像头拍照出来的图片.
                                             //关闭摄像头
           // videoSourcePlayer1.Stop();

            textBox1.Text = path + "\\" + fileName;
            pictureBox1.Image = Image.FromFile(path + "\\" + fileName, false);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            ///GeneralBasicDemo(textBox1.Text);
            */
        }
        /// <summary>
        /// 写入WORD文件或TXT文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            Object Nothing = System.Reflection.Missing.Value;
            object filename = "d://myfile.doc";
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            WordApp.Visible = false;
            Document WordDoc = WordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
           ///插入一段文本
            Microsoft.Office.Interop.Word.Paragraph para1;
            para1 = WordDoc.Content.Paragraphs.Add(ref Nothing);
            para1.Range.Text = richTextBox1.Text;
            para1.Range.InsertParagraphAfter();
            ///存盘并退出
            WordDoc.SaveAs(ref filename);
            WordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            WordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            ///写入文本文件
          ///  File.WriteAllText("d:/zhq.txt", richTextBox1.Text, Encoding.Default);
         }
        /// <summary>
        /// 位图格式转BYTE[]格式，用于在内存中直接识别
        /// </summary>
        /// <param name="bitmap"></param>
        /// <returns></returns>
        public static byte[] BitmapToGrayByte(Bitmap bitmap)
        {
            // 1.先将BitMap转成内存流
            MemoryStream ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Bmp);
            ms.Seek(0, SeekOrigin.Begin);
            // 2.再将内存流转成byte[]并返回
            byte[] bytes = new byte[ms.Length];
            ms.Read(bytes, 0, bytes.Length);
            ms.Dispose();
            return bytes;
        }
        private static void mySpeak(string s)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Rate = 2;

            speech.Speak(s);

        }
        private void button6_Click(object sender, EventArgs e)
        {
            

            IEnumerable<DataGridViewRow> enumerableList = this.dataGridView1.Rows.Cast<DataGridViewRow>();
            List<DataGridViewRow> list = (from item in enumerableList
             where item.Cells[4].Value.ToString().IndexOf("四川") >= 0
           select item).ToList();
            int matchedRowIndex = 0;
            string voice = "";
            if (list.Count > 0)
            {
                
                for (int ind = 0; ind < list.Count; ind++)
                {
                    matchedRowIndex = list[ind].Index;
                    this.dataGridView1.Rows[matchedRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
                    voice += dataGridView1.Rows[matchedRowIndex].Cells[4].Value.ToString();
                }
            }
                 Thread thread = new Thread(()=>mySpeak(voice));
                 thread.IsBackground = true;
                 thread.Start();       

        }
       
        private void button7_Click(object sender, EventArgs e)
        {

            /* Microsoft.Office.Interop.Excel.Application myapp = new Microsoft.Office.Interop.Excel.Application();
             Microsoft.Office.Interop.Excel.Workbook xBook = myapp.Workbooks.Add(Missing.Value);
             ///Microsoft.Office.Interop.Excel.Workbook xBook = myapp.Workbooks.Open("zhq.xlsx",
             ///Missing.Value, Missing.Value, Missing.Value, Missing.Value,
             /// Missing.Value, Missing.Value, Missing.Value, Missing.Value,
             /// Missing.Value, Missing.Value, Missing.Value, Missing.Value);

             Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)myapp.ActiveSheet;
          
             for (int r = 0; r < dataGridView1.Rows.Count; r++)
             {
                 for (int i = 0; i < dataGridView1.ColumnCount; i++)
                 {
                     workSheet.Cells[r +1, i + 1] = dataGridView1.Rows[r].Cells[i].Value;
                 }
                 System.Windows.Forms.Application.DoEvents();
             }
             workSheet.Columns.EntireColumn.AutoFit();//列宽自适应
             xBook.SaveAs(@"d:\test.xlsx",Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
             Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
             xBook.Close();
            */
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
            saveFileDialog.ShowDialog();
             Stream myStream = saveFileDialog.OpenFile();
             StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding("gb2312"));

           ///gridview内的空值，千万别用tostring()，否则报错！
            for (int j = 0; j <dataGridView1.Rows.Count; j++)
                {
                    string tempStr = null;
                    for (int k = 0; k <dataGridView1.Columns.Count; k++)
                    {
                    if (k > 0)
                    {
                        tempStr += "\t";
                    }
                    tempStr +=dataGridView1.Rows[j].Cells[k].Value;

                    }
                    sw.WriteLine(tempStr);
                }
                sw.Close();
               myStream.Close();
            
           MessageBox.Show("导出成功");
        }

        private void Timer_TimesUp(object sender, System.Timers.ElapsedEventArgs e)
        {
            button4_Click(sender, e);

        }
        private void button9_Click(object sender, EventArgs e)
        {
            t.Stop();
            System.Diagnostics.Debug.WriteLine("未到指定时间5秒提前终结！！！");
            button8.Enabled = true;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            t.Enabled = true; //是否触发Elapsed事件
            t.Start();
            button8.Enabled = false;
            button9.Enabled = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

   
    }
    /// <summary>
    /// 警示信息专用类
    /// </summary>
    public class xinxi
    {
        public string n0 { get; set; }
        public string n1 { get; set; }
        public string n2 { get; set; }
        public string n3 { get; set; }

        public string n4 { get; set; }
        public string n7 { get; set; }
        public string n8 { get; set; }
        public string n9 { get; set; }
    }


}
