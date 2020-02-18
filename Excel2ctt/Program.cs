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
    /// <summary>
    ///      Chứa thông tin về 1 SV/1 đối tượng cần nhập điểm
    /// </summary>
    class StudentInformation
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
    ///     Chứa các hàm để tạo nên dữ liêu về 1 đối tượng SV/đối tượng cần nhập điểm
    /// </summary>
    class StudentInformationBuilder: StudentInformation
    {
        /// <summary>
        ///     Tên/ý nghĩa của các cột thông tin đâu vào
        /// </summary>
        public static string[] FieldNames; 

        /// <summary>
        ///      Nhập điểm
        /// </summary>
        /// <param name="values"> mảng chứa thông tin điểm, tuân theo đúng trình tự như <see cref="FieldNames"/> </param>
        public StudentInformationBuilder(string[] values)
        {
            int i;
            for (i=0; i < values.Length; i++)
            {
                if (FieldNames[i]=="<diem>")
                {
                    Grade = Convert.ToDouble(values[i]);
                }
                else if (FieldNames[i] == "<stt>")
                {
                    STT = Convert.ToInt32(values[i]);
                }
                else if (FieldNames[i] == "<ten>")
                {
                    Name = values[i].Trim();
                }
                else if (FieldNames[i] == "<mssv>")
                {
                    MSSV = Convert.ToInt32(values[i]);
                }
            }
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
            Thread.Sleep(Properties.Settings.Default.WAITFORSWITCHINGAPP * 1000);
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
            int i;
            
            string[] fields;   //để phân tách các cột trong dòng thông tin

            // Phân tích chuỗi cú pháp để biết ý nghĩa của các cột thông tin. Ví dụ   <name>;<grade>
            StudentInformationBuilder.FieldNames = Properties.Settings.Default.INPUTFORMAT.Split(Properties.Settings.Default.SEPERATOR.ToCharArray());

            /// Đưa điểm vào danh sách
            try
            {
                // đọc tất cả các dòng của file đầu vào
                string[] lines = File.ReadAllLines(Properties.Settings.Default.INPUTFILENAME);
                foreach (string line in lines)
                {
                    if (line == String.Empty)
                    {
                        // nếu dòng trống, tức là SV tương ứng bị 0
                        new StudentInformation(0, 0, "", 0);
                    }
                    else
                    {
                        // tách 1 dòng thành các trường, phân tách bởi các kí tự trong cấu hình 
                        fields = line.Split(Properties.Settings.Default.SEPERATOR.ToCharArray());
                        // đưa vào cấu trúc để bóc tách và tự bổ sung vào danh sách
                        new StudentInformationBuilder(fields);
                    }
                    
                }
                DialogResult res = MessageBox.Show("Có " + lines.Length + " sinh viên. \nSau khi bấm OK, bạn có " + Properties.Settings.Default.WAITFORSWITCHINGAPP +  " giây để chuyển sang website ctt-sis và đặt con trỏ chuột vào ô điểm đầu tiên cần nhập. Để bỏ qua và kết thúc, bấm Cancel.", "Chuẩn bị", MessageBoxButtons.OKCancel);
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
