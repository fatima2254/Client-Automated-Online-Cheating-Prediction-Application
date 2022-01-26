using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CloudinaryDotNet;
using CloudinaryDotNet.Actions;
using ExcelDataReader;

namespace Admin_SSDA
{
    public partial class studentData : Form
    {
        public static Cloudinary cloudinary;
        private int _tick;
        public studentData()
        {
            InitializeComponent();
        }

      
        private void button1_Click(object sender, EventArgs e)
        {
            string CLOUD_NAME = "dhmzxoqi4";
            string API_KEY = "124261996441551";
            string API_SECRET = "TBtx4XrlE29CWvhRoJ3whJ0kNyc";

            Account account = new Account(CLOUD_NAME, API_KEY, API_SECRET);
            cloudinary = new Cloudinary(account);
            var path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            Console.WriteLine(path);
            uploadFile(path + "\\DOC.txt", GlobalConfig.UserID+ "-"+ GlobalConfig.RollNumber, "Doc");
            uploadFile(path + "\\log.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "log");
            Application.Exit();
        }
        public static void uploadFile(string path, string subfolder, string public_id)
        {
            try
            {
                var uploadParams = new RawUploadParams()
                {
                    File = new FileDescription(path),
                    PublicId = public_id,
                    Folder = "studentLogs/" + subfolder
                };
                var uploadResult = cloudinary.Upload(uploadParams); 
                 
                Console.WriteLine(uploadResult.Uri);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            _tick++;
            this.Text = _tick.ToString();
            if (_tick == 900)
            {
                t.Stop();
                timer1.Enabled = false;
                this.Text = "Done";

                MessageBox.Show("Time's Up");

                string CLOUD_NAME = "dhmzxoqi4";
                string API_KEY = "124261996441551";
                string API_SECRET = "TBtx4XrlE29CWvhRoJ3whJ0kNyc";

                Account account = new Account(CLOUD_NAME, API_KEY, API_SECRET);
                cloudinary = new Cloudinary(account);
                var path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
                Console.WriteLine(path);
                uploadFile(path + "\\DOC.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "Doc");
                uploadFile(path + "\\log.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "log");
              
                Application.Exit();
            }
        }
        System.Timers.Timer t;
        int h, m, s;
        public void filereader(string filepath)
        {
            var stream = File.Open(filepath, FileMode.Open, FileAccess.Read);



            var client = new HttpClient();
            var response =  client.GetAsync(@"https://res.cloudinary.com/dhmzxoqi4/raw/upload/v1635764665/Questions_File.xlsx");
            var streammmm = response.Result.Content.ReadAsStreamAsync();

            var reader =  ExcelReaderFactory.CreateReader(streammmm.Result);
            var result = reader.AsDataSet();
            var dtexcel = result.Tables.Cast<DataTable>();
            foreach (var row in dtexcel)
            {
                label1.Text = row.Rows[1].ItemArray[0].ToString();
                radioButton1.Text = row.Rows[1].ItemArray[1].ToString();
                radioButton2.Text = row.Rows[1].ItemArray[2].ToString();
                radioButton3.Text = row.Rows[1].ItemArray[3].ToString();
                radioButton4.Text = row.Rows[1].ItemArray[4].ToString();

                label2.Text = row.Rows[2].ItemArray[0].ToString();
                radioButton5.Text = row.Rows[2].ItemArray[1].ToString();
                radioButton6.Text = row.Rows[2].ItemArray[2].ToString();
                radioButton7.Text = row.Rows[2].ItemArray[3].ToString();
                radioButton8.Text = row.Rows[2].ItemArray[4].ToString();

                label3.Text = row.Rows[3].ItemArray[0].ToString();
                radioButton9.Text =  row.Rows[3].ItemArray[1].ToString();
                radioButton10.Text = row.Rows[3].ItemArray[2].ToString();
                radioButton11.Text = row.Rows[3].ItemArray[3].ToString();
                radioButton12.Text = row.Rows[3].ItemArray[4].ToString();

                label4.Text = row.Rows[4].ItemArray[0].ToString();
                radioButton13.Text = row.Rows[4].ItemArray[1].ToString();
                radioButton14.Text = row.Rows[4].ItemArray[2].ToString();
                radioButton15.Text = row.Rows[4].ItemArray[3].ToString();
                radioButton16.Text = row.Rows[4].ItemArray[4].ToString();

                label5.Text = row.Rows[5].ItemArray[0].ToString();
                radioButton17.Text = row.Rows[5].ItemArray[1].ToString();
                radioButton18.Text = row.Rows[5].ItemArray[2].ToString();
                radioButton19.Text = row.Rows[5].ItemArray[3].ToString();
                radioButton20.Text = row.Rows[5].ItemArray[4].ToString();

                label6.Text = row.Rows[6].ItemArray[0].ToString();
                radioButton21.Text = row.Rows[6].ItemArray[1].ToString();
                radioButton22.Text = row.Rows[6].ItemArray[2].ToString();
                radioButton23.Text = row.Rows[6].ItemArray[3].ToString();
                radioButton24.Text = row.Rows[6].ItemArray[4].ToString();

                label7.Text = row.Rows[7].ItemArray[0].ToString();
                radioButton25.Text = row.Rows[7].ItemArray[1].ToString();
                radioButton26.Text = row.Rows[7].ItemArray[2].ToString();
                radioButton27.Text = row.Rows[7].ItemArray[3].ToString();
                radioButton28.Text = row.Rows[7].ItemArray[4].ToString();

                label8.Text = row.Rows[8].ItemArray[0].ToString();
                radioButton29.Text = row.Rows[8].ItemArray[1].ToString();
                radioButton30.Text = row.Rows[8].ItemArray[2].ToString();
                radioButton31.Text = row.Rows[8].ItemArray[3].ToString();
                radioButton32.Text = row.Rows[8].ItemArray[4].ToString();

                label9.Text = row.Rows[9].ItemArray[0].ToString();
                radioButton33.Text = row.Rows[9].ItemArray[1].ToString();
                radioButton34.Text = row.Rows[9].ItemArray[2].ToString();
                radioButton35.Text = row.Rows[9].ItemArray[3].ToString();
                radioButton36.Text = row.Rows[9].ItemArray[4].ToString();

                label10.Text = row.Rows[10].ItemArray[0].ToString();
                radioButton37.Text = row.Rows[10].ItemArray[1].ToString();
                radioButton38.Text = row.Rows[10].ItemArray[2].ToString();
                radioButton39.Text = row.Rows[10].ItemArray[3].ToString();
                radioButton40.Text = row.Rows[10].ItemArray[4].ToString();

                label11.Text = row.Rows[11].ItemArray[0].ToString();
                radioButton41.Text = row.Rows[11].ItemArray[1].ToString();
                radioButton42.Text = row.Rows[11].ItemArray[2].ToString();
                radioButton43.Text = row.Rows[11].ItemArray[3].ToString();
                radioButton44.Text = row.Rows[11].ItemArray[4].ToString();

                label12.Text = row.Rows[12].ItemArray[0].ToString();
                radioButton45.Text = row.Rows[12].ItemArray[1].ToString();
                radioButton46.Text = row.Rows[12].ItemArray[2].ToString();
                radioButton47.Text = row.Rows[12].ItemArray[3].ToString();
                radioButton48.Text = row.Rows[12].ItemArray[4].ToString();

                label13.Text = row.Rows[13].ItemArray[0].ToString();
                radioButton49.Text = row.Rows[13].ItemArray[1].ToString();
                radioButton50.Text = row.Rows[13].ItemArray[2].ToString();
                radioButton51.Text = row.Rows[13].ItemArray[3].ToString();
                radioButton52.Text = row.Rows[13].ItemArray[4].ToString();

                label14.Text = row.Rows[14].ItemArray[0].ToString();
                radioButton53.Text = row.Rows[14].ItemArray[1].ToString();
                radioButton54.Text = row.Rows[14].ItemArray[2].ToString();
                radioButton55.Text = row.Rows[14].ItemArray[3].ToString();
                radioButton56.Text = row.Rows[14].ItemArray[4].ToString();

                label15.Text = row.Rows[15].ItemArray[0].ToString();
                radioButton57.Text = row.Rows[15].ItemArray[1].ToString();
                radioButton58.Text = row.Rows[15].ItemArray[2].ToString();
                radioButton59.Text = row.Rows[15].ItemArray[3].ToString();
                radioButton60.Text = row.Rows[15].ItemArray[4].ToString();
                //i++;

            }
        }
        private void studentData_Load(object sender, EventArgs e)
        {
            
            t = new System.Timers.Timer();
            t.Interval = 1000;
            t.Elapsed += OnTimeEvent;

            this.timer1.Enabled = true;
            t.Start();
            timer1.Start();
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\" + "Questions.xlsx";//@"https://res.cloudinary.com/dhmzxoqi4/raw/upload/v1635764665/Questions_File.xlsx";//
            //string path = @"C:\Users\Nabeel\Desktop\Questions - Copy.xlsx";
            filereader(path);
            //StreamReader stream = new StreamReader(path);
            //string filedata = stream.ReadToEnd();
            //var filename = Path.GetFileName(path);
            //ReadExcel(path, filename.Split('.')[1].ToString());
            ////richTextBox1.Text = filedata.ToString();
            //stream.Close();

            

            
        }
       private void studentData_FormClosing(object sender, FormClosingEventArgs e)
        {
            string CLOUD_NAME = "dhmzxoqi4";
            string API_KEY = "124261996441551";
            string API_SECRET = "TBtx4XrlE29CWvhRoJ3whJ0kNyc";


            string message = "Do you want to submit exam";
            string title = "Confirmation";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                Account account = new Account(CLOUD_NAME, API_KEY, API_SECRET);
                cloudinary = new Cloudinary(account);
                var path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
                Console.WriteLine(path);
                uploadFile(path + "\\DOC.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "Doc");
                uploadFile(path + "\\log.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "log");
                Application.Exit();
            }
            else if (result == DialogResult.No)
            {
                e.Cancel = true;

            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            string CLOUD_NAME = "dhmzxoqi4";
            string API_KEY = "124261996441551";
            string API_SECRET = "TBtx4XrlE29CWvhRoJ3whJ0kNyc";

            
            string message = "Do you want to submit exam";
            string title = "Confirmation";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                Account account = new Account(CLOUD_NAME, API_KEY, API_SECRET);
                cloudinary = new Cloudinary(account);
                var path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
                Console.WriteLine(path);
                uploadFile(path + "\\DOC.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "Doc");
                uploadFile(path + "\\log.txt", GlobalConfig.UserID + "-" + GlobalConfig.RollNumber, "log");
                Application.Exit();
            }
            else if (result == DialogResult.No)
            {
                
            }
            
        }

        private void OnTimeEvent(object sender, System.Timers.ElapsedEventArgs e)
        {try
            {
                Invoke(new Action(() =>
                {

                    s += 1;
                    if (s == 60)
                    {
                        s = 0;
                        m += 1;
                    }
                    if (m == 60)
                    {
                        m = 0;
                        h += 1;
                    }

                    timerbox.Text = string.Format("{0}:{1}:{2}", h.ToString().PadLeft(2, '0'), m.ToString().PadLeft(2, '0'), s.ToString().PadLeft(2, '0'));

                }));
            }
            catch(IOException)
            {


            }

        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            //if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
            //    conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"C:\Users\Nabeel\Desktop\Questions.xlsx" + ";Extended Properties=Excel 12.0;";//for below excel 2007
            //else
            //    conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"C:\Users\Nabeel\Desktop\Questions.xlsx" + ";Extended Properties=Excel 12.0;";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con);//here we read data from sheet1
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                    
                    foreach (DataRow row in dtexcel.Rows)
                    {
                        label1.Text = row.ItemArray[0].ToString();
                        radioButton1.Text = row.ItemArray[1].ToString();
                        radioButton2.Text = row.ItemArray[2].ToString();
                        radioButton3.Text = row.ItemArray[3].ToString();
                        radioButton4.Text = row.ItemArray[4].ToString();

                        label2.Text = row.ItemArray[0].ToString();
                        radioButton5.Text = row.ItemArray[1].ToString();
                        radioButton6.Text = row.ItemArray[2].ToString();
                        radioButton7.Text = row.ItemArray[3].ToString();
                        radioButton8.Text = row.ItemArray[4].ToString();

                        label3.Text = row.ItemArray[0].ToString();
                        radioButton9.Text = row.ItemArray[1].ToString();
                        radioButton10.Text = row.ItemArray[2].ToString();
                        radioButton11.Text = row.ItemArray[3].ToString();
                        radioButton12.Text = row.ItemArray[4].ToString();

                        label4.Text = row.ItemArray[0].ToString();
                        radioButton13.Text = row.ItemArray[1].ToString();
                        radioButton14.Text = row.ItemArray[2].ToString();
                        radioButton15.Text = row.ItemArray[3].ToString();
                        radioButton16.Text = row.ItemArray[4].ToString();

                        label5.Text = row.ItemArray[0].ToString();
                        radioButton17.Text = row.ItemArray[1].ToString();
                        radioButton18.Text = row.ItemArray[2].ToString();
                        radioButton19.Text = row.ItemArray[3].ToString();
                        radioButton20.Text = row.ItemArray[4].ToString();

                        label6.Text = row.ItemArray[0].ToString();
                        radioButton21.Text = row.ItemArray[1].ToString();
                        radioButton22.Text = row.ItemArray[2].ToString();
                        radioButton23.Text = row.ItemArray[3].ToString();
                        radioButton24.Text = row.ItemArray[4].ToString();

                        label7.Text = row.ItemArray[0].ToString();
                        radioButton25.Text = row.ItemArray[1].ToString();
                        radioButton26.Text = row.ItemArray[2].ToString();
                        radioButton27.Text = row.ItemArray[3].ToString();
                        radioButton28.Text = row.ItemArray[4].ToString();

                        label8.Text = row.ItemArray[0].ToString();
                        radioButton29.Text = row.ItemArray[1].ToString();
                        radioButton30.Text = row.ItemArray[2].ToString();
                        radioButton31.Text = row.ItemArray[3].ToString();
                        radioButton32.Text = row.ItemArray[4].ToString();

                        label9.Text = row.ItemArray[0].ToString();
                        radioButton33.Text = row.ItemArray[1].ToString();
                        radioButton34.Text = row.ItemArray[2].ToString();
                        radioButton35.Text = row.ItemArray[3].ToString();
                        radioButton36.Text = row.ItemArray[4].ToString();

                        label10.Text = row.ItemArray[0].ToString();
                        radioButton37.Text = row.ItemArray[1].ToString();
                        radioButton38.Text = row.ItemArray[2].ToString();
                        radioButton39.Text = row.ItemArray[3].ToString();
                        radioButton40.Text = row.ItemArray[4].ToString();

                        label11.Text = row.ItemArray[0].ToString();
                        radioButton41.Text = row.ItemArray[1].ToString();
                        radioButton42.Text = row.ItemArray[2].ToString();
                        radioButton43.Text = row.ItemArray[3].ToString();
                        radioButton44.Text = row.ItemArray[4].ToString();

                        label12.Text = row.ItemArray[0].ToString();
                        radioButton45.Text = row.ItemArray[1].ToString();
                        radioButton46.Text = row.ItemArray[2].ToString();
                        radioButton47.Text = row.ItemArray[3].ToString();
                        radioButton48.Text = row.ItemArray[4].ToString();

                        label13.Text = row.ItemArray[0].ToString();
                        radioButton49.Text = row.ItemArray[1].ToString();
                        radioButton51.Text = row.ItemArray[2].ToString();
                        radioButton52.Text = row.ItemArray[3].ToString();
                        radioButton53.Text = row.ItemArray[4].ToString();

                        label14.Text = row.ItemArray[0].ToString();
                        radioButton53.Text = row.ItemArray[1].ToString();
                        radioButton54.Text = row.ItemArray[2].ToString();
                        radioButton55.Text = row.ItemArray[3].ToString();
                        radioButton56.Text = row.ItemArray[4].ToString();

                        label15.Text = row.ItemArray[0].ToString();
                        radioButton57.Text = row.ItemArray[1].ToString();
                        radioButton58.Text = row.ItemArray[2].ToString();
                        radioButton59.Text = row.ItemArray[3].ToString();
                        radioButton60.Text = row.ItemArray[4].ToString();
                        //i++;

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }
    }
    
}
