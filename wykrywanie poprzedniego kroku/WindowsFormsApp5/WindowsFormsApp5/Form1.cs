using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using EasyModbus;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;

namespace WindowsFormsApp5
{
    

    public partial class Form1 : Form
    {
        SerialPort port;
        string lineReadIn;
        


        // this will prevent cross-threading between the serial port
        // received data thread & the display of that data on the central thread
        private delegate void preventCrossThreading(string x);
        private preventCrossThreading accessControlFromCentralThread;




        public Form1()
        {
            InitializeComponent();
            Application.ApplicationExit += new EventHandler(OnApplicationExit);
            WindowState = FormWindowState.Maximized;
            textBox1.Select();
            textBox1.KeyDown += textBox1_KeyDown;



            // create and open the serial port (configured to my machine)
            // this is a Down-n-Dirty mechanism devoid of try-catch blocks and
            // other niceties associated with polite programming
            const string com = "COM2";
            port = new SerialPort(com, 115200, Parity.Even, 8, StopBits.One);

            //   port.ErrorReceived += new SerialErrorReceivedEventHandler();
            try
            {
                port.Open();
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("Error: Port " + com + " jest zajęty");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Uart exception: " + ex);
            }


 

            if (port.IsOpen)
            {
                // set the 'invoke' delegate and attach the 'receive-data' function
                // to the serial port 'receive' event.

                accessControlFromCentralThread = displayTextReadIn;
                port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
                
            }

        }




        public bool ControlInvokeRequired(Control c, Action a)
        {
            if (c.InvokeRequired) c.Invoke(new MethodInvoker(delegate { a(); }));
            else return false;

            return true;
        }
        public void UpdateControl(Control myControl, Color c, String s, bool widzialnosc)
        {
            //Check if invoke requied if so return - as i will be recalled in correct thread
            if (ControlInvokeRequired(myControl, () => UpdateControl(myControl, c, s, widzialnosc))) return;
            myControl.Text = s;
            myControl.BackColor = c;
            myControl.Visible = widzialnosc;
        }












        public void OnApplicationExit(object sender, EventArgs e)
        {
            try
            {
                port.Write("LOFF\r");
            }
            catch
            {
                MessageBox.Show("Brak możlowości wyzwolenia skaneru", "Info", MessageBoxButtons.OK);
            }
            System.Windows.Forms.Application.Exit();

        }



        string wydruk;
//        string[] result = new string[100];

        // this is called when the serial port has receive-data for us.
        private void port_DataReceived(object sender, SerialDataReceivedEventArgs rcvdData)
        {
            start = DateTime.Now;

            while (port.BytesToRead > 0)
            {

                //   lineReadIn += port.ReadExisting();

                lineReadIn += port.ReadExisting();
                //   lineReadIn += Environment.NewLine;

                // lineReadIn += "\r\n";
                //   lineReadIn += lineReadIn;

           //     flaga = false;
           //     aTimer.Stop();

                Thread.Sleep(25);
            }

            //   flaga = true;


            // display what we've acquired.
            
            UpdateControl(label2, SystemColors.Control, "Błąd czytania barkodu.", false);

            lineReadIn = lineReadIn.ToUpper();
            string firstletter = lineReadIn.Remove(1);

            lineReadIn = Regex.Replace(lineReadIn, @"\s+", string.Empty);
            if (lineReadIn.Length > 19  || lineReadIn.Equals("E"))
            {
                UpdateControl(label2, Color.Red, "Błąd czytania barkodu.", true);
                this.BackColor = Color.FromArgb(255, 128, 128);
            }
            else if (lineReadIn.Length < 19)
            {
                UpdateControl(label2, Color.Red, "Błąd czytania barkodu.", true);
                this.BackColor = Color.FromArgb(255, 128, 128);
            }
            else if (firstletter.Equals("A"))
            {
                tworzeniepliku(lineReadIn);
                displayTextReadIn(lineReadIn);

                Task.Run(async () => await Test());
                UpdateControl(label3, Color.LawnGreen, "Krok OK", false);
                this.BackColor = SystemColors.Control;
            }

            displayTextReadIn(lineReadIn);
            wydruk = lineReadIn;
            lineReadIn = string.Empty;




        }// end function 'port_dataReceived'


        private async Task Test()
        {

            await Task.Delay(5000);
            UpdateControl(label3, Color.LawnGreen, "Krok OK", false);
            this.BackColor = SystemColors.Control;
            // await - jakas inna dluga operacja
        }






        // this, hopefully, will prevent cross threading.
        private void displayTextReadIn(string ToBeDisplayed)          //wyswietlanie sygnalu na drugim texboxie
        {
            if (textBox1.InvokeRequired)
                textBox1.BeginInvoke(accessControlFromCentralThread, ToBeDisplayed);
            else
                textBox1.Text = ToBeDisplayed;
          
        }


        DateTime stop, start;
        //---------------------------------------------------------------------------------------------

        



        private int tworzeniepliku(string sn)
        {
            sn = Regex.Replace(sn, @"\s+", string.Empty);

            //     if(sn.Length > 8)
            //      sn = sn.Remove(8);

            if (sn == "ERROR" || sn.Length != 19)
            {
                UpdateControl(label2, Color.Red, "Błąd czytania barkodu.", true);
                UpdateControl(label3, Color.LawnGreen, "Krok OK", false);
                this.BackColor = Color.FromArgb(255, 128, 128);
                return 0;
            }
            else
            {
                this.BackColor = Color.FromArgb(0, 192, 0);
                UpdateControl(label3, Color.LawnGreen, "Krok OK", true); 
            }
            stop = DateTime.Now;
            string stop_String = stop.ToString("yyyy-MM-dd HH:mm:ss");
            // textBox1.Text = sn;

            string sciezka = (@"C:/tars/");      //definiowanieścieżki do której zapisywane logi
            string sourceFile = @"C:/tars/" + @sn + @"(" + @stop.ToString("yyyy-MM-dd HH-mm-ss") + @")" + @".Tars";
            string destinationFile = @"C:/copylogi/" + @stop.Day + @"-" + @stop.Month + @"-" + @stop.Year + @"/" + @sn + @"(" + @stop.ToString("yyyy-MM-dd HH-mm-ss") + @")" + @".Tars";

            stop = DateTime.Now;


            if (Directory.Exists(sciezka))       //sprawdzanie czy sciezka istnieje
            {
                ;
            }
            else
                System.IO.Directory.CreateDirectory(sciezka); //jeśli nie to ją tworzy

            if (Directory.Exists(@"C:/copylogi/" + @stop.Day + @"-" + @stop.Month + @"-" + @stop.Year + @"/"))       //sprawdzanie czy sciezka istnieje
            {
                ;
            }
            else
                System.IO.Directory.CreateDirectory(@"C:/copylogi/" + @stop.Day + @"-" + @stop.Month + @"-" + @stop.Year + @"/"); //jeśli nie to ją tworzy


            try
            {
                using (StreamWriter sw = new StreamWriter("C:/tars/" + sn + "(" + @stop.ToString("yyyy-MM-dd HH-mm-ss") + ")" + ".Tars"))
                {


                    sw.WriteLine("S{0}", sn);
                    sw.WriteLine("CITRON");
                    sw.WriteLine("NPLKWIM0T26B2PR3");
                    sw.WriteLine("PPRESS");
                    sw.WriteLine("Ooperator");


                    // sw.WriteLine("[" + start.Year + "-" + stop.Month + "-" + stop.Day + " " + stop.Hour + ":" + stop.Minute + ":" + stop.Second);
                    sw.WriteLine("[" + start.ToString("yyyy-MM-dd HH:mm:ss"));
                    sw.WriteLine("]" + stop_String);


                    sw.WriteLine("TP");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                port.Write("LOFF\r");
            }

            try
            {
                File.Copy(sourceFile, destinationFile, true);
            }
            catch (IOException iox)
            {
                MessageBox.Show(iox.Message);
            }


            return 1;


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                port.Write("LON\r");
            }
            catch
            {
                MessageBox.Show("Brak możlowości wyzwolenia skaneru", "Info", MessageBoxButtons.OK);
            }
          //  port.WriteLine("LON");
            //port.Write("4C4F4E");
            //port.WriteLine("4C4F4E");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                port.Write("LOFF\r");
            }
            catch
            {
                MessageBox.Show("Brak możlowości wyzwolenia skaneru", "Info", MessageBoxButtons.OK);
            }
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        string firstletter;

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            
            // Determine whether the key entered is the F1 key. If it is, display Help.
            if (e.KeyCode == Keys.Enter)
            {

                if (sprawdzeniekrok(textBox1.Text) == 1)
                {
                    MessageBox.Show("GIT");
                }
                else
                    MessageBox.Show("WUJO");

                    UpdateControl(label2, SystemColors.Control, "Błąd barkodu.", false);
                textBox1.Text = textBox1.Text.ToUpper();
                if(textBox1.Text.Length > 1)
                    firstletter = textBox1.Text.Remove(1);

                textBox1.Text = Regex.Replace(textBox1.Text, @"\s+", string.Empty);
                if (textBox1.Text.Length > 19 || textBox1.Text.Equals("E"))
                {
                    UpdateControl(label2, Color.Red, "Błąd barkodu.", true);
                    this.BackColor = Color.FromArgb(255, 128, 128);
                }
                else if (textBox1.Text.Length < 19)
                {
                    UpdateControl(label2, Color.Red, "Błąd barkodu.", true);
                    this.BackColor = Color.FromArgb(255, 128, 128);
                }
                else if (firstletter.Equals("A"))
                {
    //                this.BackColor = Color.FromArgb(0, 192, 0);
   //                 UpdateControl(label3, Color.LawnGreen, "Krok OK", true);

                    tworzeniepliku(textBox1.Text);
                    displayTextReadIn(textBox1.Text);

                    Task.Run(async () => await Test());
                  //  UpdateControl(label3, Color.LawnGreen, "Krok OK", false);
                  //  this.BackColor = SystemColors.Control;
                }
             //   textBox1.Text = string.Empty;


            }
        }


        const int M_NIENARODZONY = 1;
        const int M_BRAK_KROKU = 2;
        const int M_FAIL = 3;
        const int M_BRAK_POLACZENIA_Z_MES = 4;


        public int sprawdzanieMES(string SerialTxt)
        {
            using (MESwebservice.BoardsSoapClient wsMES = new MESwebservice.BoardsSoapClient("BoardsSoap"))
            {
                DataSet Result;
                try
                {
                    Result = wsMES.GetBoardHistoryDS(@"itron", SerialTxt);
                }
                catch
                {
                    return M_BRAK_POLACZENIA_Z_MES;
                }

                var Test = Result.Tables[0].TableName;
                if (Test != "BoardHistory") return M_NIENARODZONY; //numer produktu nie widnieje w systemie MES

                //where row.Field<string>("Test_Process").ToUpper() == "FVT / HOT_PRESS".ToUpper() || row.Field<string>("Test_Process").ToUpper() == "FVT / HOT_PRESS".ToUpper()
                var data = (from row in Result.Tables["BoardHistory"].AsEnumerable()
                            where row.Field<string>("Test_Process").ToUpper() == "QC / BB_GRN".ToUpper() || row.Field<string>("Test_Process").ToUpper() == "QC / BB_GRN".ToUpper()
                            select new
                            {
                                TestProcess = row.Field<string>("Test_Process"),
                                TestType = row.Field<string>("TestType"),
                                TestStatus = row.Field<string>("TestStatus"),
                                StartDateTime = row.Field<DateTime>("StartDateTime"),
                                StopDateTime = row.Field<DateTime>("StopDateTime"),
                            }).FirstOrDefault();


                if (data != null)
                {
                    //sprawdzamy PASS w poprzednim kroku
                    if ("PASS" == data.TestStatus.ToUpper()) return 0; //wszystko jest OK
                    else return M_FAIL;
                }
                else return M_BRAK_KROKU; //brak poprzedniego kroku
            }
        }




        private int sprawdzeniekrok(string sn)
        {
            int Result;

            Result = sprawdzanieMES(sn); //przykladowy numer seryjny 9100000668
            switch (Result)
            {
                case M_BRAK_POLACZENIA_Z_MES:
                    MessageBox.Show("Brak połączenia z MES.", "Info", MessageBoxButtons.OK);
                    //label8.Text = "Brak połączenia z MES.";
                    break;

                case M_NIENARODZONY:
                    MessageBox.Show("Numer nienarodzony w MES.", "Info", MessageBoxButtons.OK);
                    //label8.Text = "Numer nienarodzony w MES.";
                    break;

                case M_BRAK_KROKU:
                    MessageBox.Show("Brak poprzedniego kroku.", "Info", MessageBoxButtons.OK);
                    // label8.Text = "Brak poprzedniego kroku.";
                    break;

                case M_FAIL:
                    MessageBox.Show("Poprzedni krok = FAIL.", "Info", MessageBoxButtons.OK);
                    //  label8.Text = "Poprzedni krok = FAIL.";
                    break;

                default:
                    //  MessageBox.Show("Wszystko jest OK", "Info", MessageBoxButtons.OK);                   
                    return 1;
            }
            return 0;
        }










    }
}
