using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data.Linq;
using System.Windows.Forms;
using System.IO;

namespace Excel2ctt
{
    class  StudentInformation
    {
        static StudentInformation()
        {
            list = new List<StudentInformation>();
        }
        public static List<StudentInformation> list;
        public int STT;
        public int MSSV;
        public string Name;
        public double Grade;

        public StudentInformation()
        {
            list.Add(this);
        }

        public StudentInformation(int STT, int MSSV, string Name, double Grade)
        {
            StudentInformation NewItem = new StudentInformation();
            NewItem.STT = STT;
            NewItem.Grade = Grade;
            NewItem.MSSV = MSSV;
            NewItem.Name = Name;
        }
    }

    /// <summary>
    ///      Nhập điểm vào website của pđt
    /// </summary>
    class CTTSite
    {
        public static List<StudentInformation> StudentList = null;

        public static void AutoTyping()
        {
            StudentInformation si;
            Thread.Sleep(2000);
            for (int i = 0; i < StudentList.Count; i++)
            {
                si = StudentList[i];
                SendKeys.SendWait(si.Grade.ToString());
                // Chuyển tới bản ghi kế tiếp
                SendKeys.SendWait(Properties.Settings.Default.GOTONEXTRECORD);
            }
        }
    }

    class Program
    {

        
        static void Main(string[] args)
        {
            /// Đưa điểm vào danh sách
            try
            {
                string[] lines = File.ReadAllLines("diem.txt");
                foreach (string line in lines)
                {
                    if (line == String.Empty)
                    {
                        new StudentInformation(0, 0, "", 0);
                    }
                    else
                    {
                        new StudentInformation(0, 0, "", Convert.ToDouble(line));
                    }
                    
                }
                DialogResult res = MessageBox.Show("Có " + lines.Length + " sinh viên. \nSau khi bấm OK, bạn có 5 giây để chuyển sang website ctt-sis và đặt con trỏ chuột vào ô điểm đầu tiên cần nhập. Để bỏ qua và kết thúc, bấm Cancel.", "Chuẩn bị", MessageBoxButtons.OKCancel);
                if (res != DialogResult.OK)
                {
                    return;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

            /// Đưa thông tin điểm vào class phụ trách
            CTTSite.StudentList = StudentInformation.list;

            /// Thực hiện nhập điểm
            CTTSite.AutoTyping();
        }
    }
}
