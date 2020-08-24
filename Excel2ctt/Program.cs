using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

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

    class Program
    {


        static void Main(string[] args)
        {

            string[] fields;   //để phân tách các cột trong dòng thông tin

            // Phân tích chuỗi cú pháp để biết ý nghĩa của các cột thông tin. Ví dụ   <name>;<grade>
            StudentInformationBuilder.FieldNames = Properties.Settings.Default.INPUTFORMAT.Split(Properties.Settings.Default.SEPERATOR.ToCharArray());

            /// Tạo đối tượng cần nhập điểm
            Victim site = new Victim(Properties.Settings.Default.AUTOACTIVATE);

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

                /// Tự động Activate site nạn nhân để tiền kiểm tra.
                site.Activate();

                /// Quay trở lại tiến trình hiện thời để nhìn rõ thông báo
                Victim.SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);

                /// Hiện thông báo confirm với người dùng
                DialogResult res = MessageBox.Show("Có " + lines.Length + " dòng nhập liệu. \nSau khi bấm OK, bạn có " + Properties.Settings.Default.WAITFORSWITCHINGAPP + " giây để chuyển sang website ctt-sis và đặt con trỏ chuột vào ô điểm đầu tiên cần nhập. Để bỏ qua và kết thúc, bấm Cancel.\n" + lines[0] ?? "" + "\n" + lines[1] ?? "" + "\n" + lines[2] ?? "", "Chuẩn bị", MessageBoxButtons.OKCancel);
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
                    Victim.StudentList = StudentInformation.list;

            do
            {
                /// Thực hiện nhập điểm
                site.AutoTyping();
                if (Properties.Settings.Default.KEEPALIVE)
                {
                    if (MessageBox.Show("Chọn lại vào textbox nhập điểm đầu tiên ở trang ctsv, rồi bấm okay để lặp lại quá trình. Bấm Cancel để kết thúc chương trình", "zz", MessageBoxButtons.OKCancel)==DialogResult.Cancel)
                    {
                        break;
                    }
                }
            } while (true);
        }
    }
}
