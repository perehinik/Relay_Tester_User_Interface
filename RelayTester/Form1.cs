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
using System.IO;



namespace RelayTester
{
    public partial class Form1 : Form
    {
        string testerPortName = null;
        string returnMessage = null;

        float CoilVoltage = 0;
        string CoilVoltageStr = null;
        float contactCurrent = 0;
        string contactCurrentStr = null;

        static int relayIndex = 0;
        static int stepIndex = 0;

        string relayCurrent = null;
        string coilVoltage = null;

        bool testActive = false;
        bool testSwitchTime = true;

        float[] CurrentOFF = new float[18];
        float[] CurrentON = new float[18];
        float[] VoltageOFF = new float[18];
        float[] VoltageON = new float[18];
        float[] RelTimeOFF = new float[18];
        float[] RelTimeON = new float[18];
        float[] ResOFF = new float[18];
        float[] ResON = new float[18];

        int testNum = 1;
        int tempTestNum = 1;

        public delegate void ClearTableDelegate();
        public delegate void AddRowDelegate(string c1, string c2, string c3, string c4, string c5, string c6, string c7, string c8, string c9);

        Bitmap bitmap = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void AddRow(string c1, string c2, string c3, string c4, string c5, string c6, string c7, string c8, string c9)
        {
            object[] RowArray = new object[] {c1,c2,c3,c4,c5,c6,c7,c8,c9};
            if (this.InvokeRequired)
                dataGridView1.Invoke((MethodInvoker)(() => dataGridView1.Rows.Add(RowArray)));
            else dataGridView1.Rows.Add(RowArray);
        }
        private void ClearTable()
        {
            if (this.InvokeRequired)
                dataGridView1.Invoke((MethodInvoker)(() => dataGridView1.Rows.Clear()));
            else dataGridView1.Rows.Clear();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "";
            comboBox1.SelectedIndex = 0;
            comboBox3.SelectedIndex = 3;
            comboBox4.SelectedIndex = 0;
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
            ConnectionChkTim.Tick += new EventHandler(OnTimeEvent);
            ConnectionCloseTim.Tick += new EventHandler(OnConCloseEvent);
            //ConnectionChkTim.Start();
            //ConnectionCloseTim.Start();

            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void OnTimeEvent(object source, EventArgs e)
        { 
            Console.WriteLine("CHKEvent");
            if (serialPort1.IsOpen) 
                try
                {
                    serialPort1.Write("Are you alive");
                }
                catch (IOException)
                {
                    TesterDisconnect();
                    label1.Text = "Relay tester disconnected from PC";
                    button1.BackColor = Color.Transparent;
                    button1.Text = "Connect";
                    button1.FlatAppearance.BorderColor = Color.Lime;
                    button2.BackColor = Color.Transparent;
                    button2.Enabled = false;
                }
        }

        private void OnConCloseEvent(object source, EventArgs e)
        {
            Console.WriteLine("OnConCloseEvent");
            if (serialPort1.IsOpen)  
                {
                    TesterDisconnect();
                    label1.Text = "Relay tester disconnected from PC";
                    button1.BackColor = Color.Transparent;
                    button1.Text = "Connect";
                    button1.FlatAppearance.BorderColor = Color.Lime;
                    button2.BackColor = Color.Transparent;
                    button2.Enabled = false;
                }
        }

        private void StartTest(int stepIndexFunc, int relIndexFunc)
        {
            ConnectionChkTim.Stop();

            if (!testSwitchTime) //if we not need to check relay switch time 
            {
                if (comboBox1.Text == "White relay PCB")
                {
                    if (relIndexFunc % 2 == 1) { relIndexFunc += 1; relayIndex = relIndexFunc; }
                    switch (stepIndexFunc)
                    {
                        case 100: serialPort1.Write("$SetCoilVoltage" + coilVoltage); break;
                        case 101: serialPort1.Write("$SetRelCurrent" + relayCurrent); break;
                        case 1: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 2:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentOFF");
                            break;
                        case 3: serialPort1.Write("@GetVoltageOFF"); break;
                        case 4: serialPort1.Write("$DIOOFF" + relIndexFunc); break;
                        case 5:
                            serialPort1.Write("$RelON" + (relIndexFunc / 2));
                            Console.WriteLine(relIndexFunc); break;
                        case 6: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 7:
                            System.Threading.Thread.Sleep(150);
                            serialPort1.Write("@GetCurrentON");
                            break;
                        case 8: serialPort1.Write("@GetVoltageON"); break;
                        case 9: serialPort1.Write("$DIOOFF" + relIndexFunc); break;
                        case 10: serialPort1.Write("$RelOFF" + (relIndexFunc / 2)); break;
                    }
                }
                else if (comboBox1.Text == "Black relay PCB")
                {
                    switch (stepIndexFunc)
                    {
                        case 100: serialPort1.Write("$SetCoilVoltage" + coilVoltage); break;
                        case 101: serialPort1.Write("$SetRelCurrent" + relayCurrent); break;
                        case 1: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 2:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentOFF");
                            break;
                        case 3: serialPort1.Write("@GetVoltageOFF"); break;
                        case 4: serialPort1.Write("$DIOOFF" + relIndexFunc); break;
                        case 5: serialPort1.Write("$RelON" + ((relIndexFunc / 2) + (relIndexFunc % 2))); break;
                        case 6: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 7:
                            System.Threading.Thread.Sleep(150);
                            serialPort1.Write("@GetCurrentON");
                            break;
                        case 8: serialPort1.Write("@GetVoltageON"); break;
                        case 9: serialPort1.Write("$DIOOFF" + relIndexFunc); break;
                        case 10: serialPort1.Write("$RelOFF" + ((relIndexFunc / 2) + (relIndexFunc % 2))); break;
                    }
                }
            }
            else //if we NEED to check relay switch time 
            {
                if (comboBox1.Text == "White relay PCB")
                {
                    if (relIndexFunc % 2 == 1) { relIndexFunc += 1; relayIndex = relIndexFunc; }
                    switch (stepIndexFunc)
                    {
                        case 100: serialPort1.Write("$SetCoilVoltage" + coilVoltage); break;
                        case 101: serialPort1.Write("$SetRelCurrent" + relayCurrent); break;
                        case 1: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 2:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentOFF");
                            break;
                        case 3: serialPort1.Write("@GetVoltageOFF"); break;
                        case 4: 
                            serialPort1.Write("@RelTimeON" + (relIndexFunc / 2));
                            Console.WriteLine(relIndexFunc); break;
                        case 5:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentON");
                            break;
                        case 6: serialPort1.Write("@GetVoltageON"); break;
                        case 7: serialPort1.Write("@RelTimeOFF" + (relIndexFunc / 2)); break;
                        case 8:
                            serialPort1.Write("$DIOOFF" + relIndexFunc); 
                            stepIndex = 10;
                            break;
                    }
                }
                else if (comboBox1.Text == "Black relay PCB")
                {
                    switch (stepIndexFunc)
                    {
                        case 100: serialPort1.Write("$SetCoilVoltage" + coilVoltage); break;
                        case 101: serialPort1.Write("$SetRelCurrent" + relayCurrent); break;
                        case 1: serialPort1.Write("$DIOON" + relIndexFunc); break;
                        case 2:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentOFF");
                            break;
                        case 3: serialPort1.Write("@GetVoltageOFF"); break;
                        case 4: serialPort1.Write("@RelTimeON" + ((relIndexFunc / 2) + (relIndexFunc % 2))); break;
                        case 5:
                            System.Threading.Thread.Sleep(50);
                            serialPort1.Write("@GetCurrentON");
                            break;
                        case 6: serialPort1.Write("@GetVoltageON"); break;
                        case 7: serialPort1.Write("@RelTimeOFF" + ((relIndexFunc / 2) + (relIndexFunc % 2))); break; 
                        case 8:
                            serialPort1.Write("$DIOOFF" + relIndexFunc);
                            stepIndex = 10;
                            break;
                    }
                }
            }
            ConnectionChkTim.Start();
        }

        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            ConnectionCloseTim.Stop();
            Console.WriteLine("Data received");
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            if (indata.Contains('@') || indata.Contains('$'))
            {
                float progressBarValue = 100 / testNum / 16 * relayIndex + ((tempTestNum-1)*16)*(100 / testNum / 16);
                if (progressBarValue <= 100) {progressBar1.Value = (int)progressBarValue;}
                else {progressBar1.Value = 100;}

                Console.WriteLine("Rel  " + relayIndex + "  Step  " + stepIndex + ' ' + indata);
                stepIndex++;        
                if (stepIndex == 102)  { stepIndex = 1; }
       
                else if (stepIndex >= 11 && relayIndex < 17 && stepIndex < 100)
                {
                    if (comboBox1.Text == "Black relay PCB") { relayIndex++; }
                    else if (comboBox1.Text == "White relay PCB") { relayIndex+=2; }
                    stepIndex = 1;
                }
                if (relayIndex >= 17)  // Test finish??
                {
                    if (tempTestNum < testNum)
                    {
                        tempTestNum++;
                        relayIndex = 1;
                        stepIndex = 1;
                        Console.WriteLine("Test Number  "+ tempTestNum);
                        System.Threading.Thread.Sleep(1000);
                        WriteResults();
                    }
                    else {
                        tempTestNum = 1;
                        Console.WriteLine("TestFinish");
                        relayIndex = 0;
                        stepIndex = 0;
                        progressBar1.Value = 100;
                        WriteResults();
                        testActive = false;
                        serialPort1.Write("$AllOFF");
                    }
                }
            }
            else if (indata.Contains("I am Tester")){ returnMessage = indata; }

            if (indata.Contains('@'))
            {
                if (indata.Contains("CurrentOFF")) {CurrentOFF[relayIndex] = StrToFloat(indata);}
                if (indata.Contains("CurrentON"))  {CurrentON[relayIndex] = StrToFloat(indata); }
                if (indata.Contains("VoltageOFF")) {VoltageOFF[relayIndex] = StrToFloat(indata);}
                if (indata.Contains("VoltageON"))  {VoltageON[relayIndex] = StrToFloat(indata); }
                if (indata.Contains("RelTimeOFF")) {RelTimeOFF[relayIndex] = StrToFloat(indata);}
                if (indata.Contains("RelTimeON"))  {RelTimeON[relayIndex] = StrToFloat(indata); }
            }

            if (relayIndex != 0 && stepIndex != 0) {StartTest(stepIndex, relayIndex);}
            ConnectionCloseTim.Start();
        }


        float StrToFloat(string Buf)
        {
            string IntMas =  "0123456789" ;
            int Len = Buf.Length;
            double TempInt = 0;
            int dotIndex = 0;

            for (int i = 0; i < Len; i++)
            {
                if (Buf[i] == '.' || Buf[i] == ',')
                {
                    dotIndex = 1;
                    i++;
                }
                for (int j = 0; j < 10; j++)
                {
                    if (Buf[i] == IntMas[j] && dotIndex == 0)
                    {
                        TempInt = (TempInt * 10) + j;
                        break;
                    }
                    else if (Buf[i] == IntMas[j] && dotIndex > 0)
                    {
                        TempInt = (Math.Pow(0.1D, dotIndex) * j) + TempInt;
                        dotIndex++;
                        break;
                    }
                }
            }
            return (float)TempInt;
        }



        public void WriteResults()
        {
            dataGridView1.Rows.Clear();
            string relayLetterIndex = "ABCDEFGH";

            if (comboBox1.Text == "Black relay PCB")
            {
                for (int i = 1; i < 17; i++)
                {
                    if (CurrentOFF[i] < 0.03) CurrentOFF[i] = 0;
                    if (CurrentON[i] < 0.03) CurrentON[i] = 0;
                    ResOFF[i] = VoltageOFF[i] / CurrentOFF[i];
                    ResON[i] = VoltageON[i] / CurrentON[i];
                    int tempRelNum = ((i / 2) + (i % 2));
                    AddRow(tempRelNum + "(" + relayLetterIndex[tempRelNum - 1] + ")." + (2-(i % 2)),
                                              CurrentOFF[i].ToString("N4"),
                                              VoltageOFF[i].ToString("N4"),
                                              ResOFF[i].ToString("N4"),
                                              CurrentON[i].ToString("N4"),
                                              VoltageON[i].ToString("N4"),
                                              ResON[i].ToString("N4"),
                                              RelTimeON[i].ToString(),
                                              RelTimeOFF[i].ToString());
                    Console.WriteLine(CurrentOFF[i] + "   " + VoltageOFF[i] + "  " + i);
                }
                bitmap = (Bitmap)Bitmap.FromFile(Application.StartupPath + @"\resources\BlackRel\BlackRelAll.png");

                for (int i = 1; i < 17; i++)
                {
                    if (ResOFF[i] < 9 || ResON[i] > 0.9 || RelTimeON[i] > 1300 || RelTimeOFF[i] > 500)
                    {
                        dataGridView1.Rows[i - 1].DefaultCellStyle.BackColor = Color.Red;
                        bitmap = CombineBitmap(bitmap, new[] { Application.StartupPath + @"\resources\BlackRel\BlackRel" + ((i / 2) + (i % 2)) + ".png" });
                    }
                    else
                    {
                        dataGridView1.Rows[i - 1].DefaultCellStyle.BackColor = Color.Green;
                    }
                }
            }
            else if (comboBox1.Text == "White relay PCB")
            {
                for (int i = 2; i < 17; i+=2)
                {
                    if (CurrentOFF[i] < 0.03) CurrentOFF[i] = 0;
                    if (CurrentON[i] < 0.03) CurrentON[i] = 0;
                    ResOFF[i] = VoltageOFF[i] / CurrentOFF[i];
                    ResON[i] = VoltageON[i] / CurrentON[i];
                    int tempRelNum = (i / 2);
                    AddRow(tempRelNum + "(" + relayLetterIndex[tempRelNum - 1] + ")",
                                              CurrentOFF[i].ToString("N4"),
                                              VoltageOFF[i].ToString("N4"),
                                              ResOFF[i].ToString("N4"),
                                              CurrentON[i].ToString("N4"),
                                              VoltageON[i].ToString("N4"),
                                              ResON[i].ToString("N4"),
                                              RelTimeON[i].ToString(),
                                              RelTimeOFF[i].ToString());
                    Console.WriteLine(CurrentOFF[i] + "   " + VoltageOFF[i] + "  " + i);
                }
                bitmap = (Bitmap)Bitmap.FromFile(Application.StartupPath + @"\resources\WhiteRel\WhiteRelAll.png");
                for (int i = 2; i < 17; i+=2)
                {
                    if (ResOFF[i] < 9 || ResON[i] > 0.9)
                    {
                        dataGridView1.Rows[(i / 2 - 1)].DefaultCellStyle.BackColor = Color.Red;
                        bitmap = CombineBitmap(bitmap, new[] { Application.StartupPath + @"\resources\WhiteRel\WhiteRel" + ((i / 2) + (i % 2)) + ".png" });
                    }
                    else
                    {
                        dataGridView1.Rows[(i / 2 - 1)].DefaultCellStyle.BackColor = Color.Green;
                    }
                }
            }
                
            pictureBox0.Image = bitmap;
        }




        private void button1_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                if (TesterConnect())
                {
                    label1.Text = "Relay tester connected to " + testerPortName;
                    button1.BackColor = Color.Lime;
                    button1.Text = "Disconnect";
                    button2.BackColor = Color.Lime;
                    button2.Enabled = true;
                }
                else
                {
                    label1.Text = "Can not detect the tester";
                }
            }
            else
            {
                if (TesterDisconnect())
                {
                    label1.Text = "Relay tester disconnected";
                    button1.BackColor = Color.Transparent;
                    button1.Text = "Connect";
                    button1.FlatAppearance.BorderColor = Color.Lime;
                }
                else
                {
                    label1.Text = "Disconnecting ERROR!!!";
                    button1.BackColor = Color.Transparent;
                    button1.Text = "Connect";
                    button1.FlatAppearance.BorderColor = Color.Lime;
                }

                button2.BackColor = Color.Transparent;
                button2.Enabled = false;
            }
        }


        private void serialPort1_PinChanged(object sender, SerialPinChangedEventArgs e)
        {
            label1.Text = "Disconnecting ERROR!!!";
        }


        public bool TesterConnect()
        {
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                Console.WriteLine("try to open port "+port);
                serialPort1.PortName = port;
                serialPort1.BaudRate = 38400;
                serialPort1.Open();  //CHECK THAT - Msg:Device is not work
                System.Threading.Thread.Sleep(100);
                if (!serialPort1.IsOpen)
                {  
                    Console.WriteLine("can not open port " + port);
                    serialPort1.Close();
                }
                else
                {
                    Console.WriteLine("port " + port +" is open, checking the device");
                    serialPort1.Write("WhoAreYou");
                    System.Threading.Thread.Sleep(600);
                    
                    Console.WriteLine("Device send -    " + returnMessage);
                    if (returnMessage=="I am Tester,Relay Tester")
                    {
                        Console.WriteLine("Tester detected on " + port);
                        testerPortName = port;
                        return true;
                    }
                    else
                    {
                        serialPort1.Close();
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            return false;
        }


        public bool TesterDisconnect()
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.Close();
                }
                catch (IOException) { return false; }
            }
            return true;
        }

        private void comboBox3_LostFocus(object sender, EventArgs e)
        {
            if (comboBox3.Text == "")  
            {
                comboBox3.Text = "12";
            }
        }
        private void comboBox4_LostFocus(object sender, EventArgs e)
        {
            if (comboBox4.Text == "")
            {
                comboBox4.Text = "1";
            }
        }
        private void groupBox2_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Tester send -  :))  ");
            this.ActiveControl = label1;
        }
        private void groupBox1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Tester send -  :))  ");
            this.ActiveControl = label1;
        }
       
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           if (comboBox1.Text == "Black relay PCB")
            {
                bitmap = (Bitmap)Bitmap.FromFile(Application.StartupPath + @"\resources\BlackRel\BlackRelAll.png");
                pictureBox0.Image = bitmap;
            }
            else if (comboBox1.Text == "White relay PCB")
            {
                bitmap = (Bitmap)Bitmap.FromFile(Application.StartupPath + @"\resources\WhiteRel\WhiteRelAll.png");
                pictureBox0.Image = bitmap;
            }
            else if(comboBox1.Text == "SPST")
            {

            }
            else if (comboBox1.Text == "SPDT")
            {

            }
            else if (comboBox1.Text == "DPST")
            {

            }
            else if (comboBox1.Text == "DPDT")
            {

            }
        }
        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            CoilVoltageStr = parsePrepare(comboBox3.Text, 1);
            float.TryParse(comboBox3.Text, out CoilVoltage);
            if (CoilVoltage > 12)
            {
                CoilVoltageStr = "12";
            }
            comboBox3.Text = CoilVoltageStr;
            comboBox3.SelectionStart = comboBox3.Text.Length;
        }
        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            string comboBox4Text = parsePrepare(comboBox4.Text, 1); 
            comboBox4.Text = comboBox4Text;
            int.TryParse(comboBox4.Text, out testNum);
            if (testNum > 100)
            {
                testNum = 100;
            }
            comboBox4.Text = ""+testNum;
            comboBox4.SelectionStart = comboBox4.Text.Length;
        }
        private string parsePrepare(string value, int numAfterPoint)
        {
            string outValue = null;
            bool comaIndex = false;
            int numAfterPointOut = 0;
            foreach (char item in value)
            {
                if ((item == '1' || item == '2' || item == '3' || item == '4' || item == '5' || item == '6' || item == '7' || item == '8' || item == '9' || item == '0') && (numAfterPointOut < numAfterPoint))
                {
                    outValue = outValue + item;
                    if (comaIndex == true)
                    {
                        numAfterPointOut++;
                    }
                }
                else if (item == ',' && comaIndex == false && numAfterPoint != 0)
                {
                    outValue = outValue + item;
                    comaIndex = true;
                }
                else if (item == '.' && comaIndex == false && numAfterPoint != 0)
                {
                    outValue = outValue + ',';
                    comaIndex = true;
                }
            }
            return outValue;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (!testActive)
            {
                testActive = true;
                stepIndex = 100;
                relayIndex = 1;
                relayCurrent = "1";
                coilVoltage = comboBox3.Text;
                testSwitchTime = checkBox1.Checked;
                StartTest(100, 1);
            }
        }
        public static Bitmap CombineBitmap(Bitmap bitmapTemp, IEnumerable<string> files)
        {
            //read all images into memory
            List<Bitmap> images = new List<Bitmap>();
            Bitmap finalImage = null;

            if (bitmapTemp != null) images.Add(bitmapTemp);

            try
            {
                int width = 0;
                int height = 0;

                foreach (string image in files)
                {
                    // create a Bitmap from the file and add it to the list
                    Bitmap bitmap = new Bitmap(image);

                    // update the size of the final bitmap
                    width += bitmap.Width;
                    height = bitmap.Height > height ? bitmap.Height : height;

                    images.Add(bitmap);
                }

                // create a bitmap to hold the combined image
                finalImage = new Bitmap(width, height);

                // get a graphics object from the image so we can draw on it
                using (Graphics g = Graphics.FromImage(finalImage))
                {
                    // set background color
                    g.Clear(Color.Transparent);

                    // go through each image and draw it on the final image
                    foreach (Bitmap image in images)
                    {
                        g.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));
                    }
                }

                return finalImage;
            }
            catch (Exception)
            {
                if (finalImage != null) finalImage.Dispose();
                throw;
            }
            finally
            {
                // clean up memory
                foreach (Bitmap image in images)
                {
                    image.Dispose();
                }
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Designed by Ivan Perehiniak\niv.perehinik@gmail.com");
        }
    }
}
