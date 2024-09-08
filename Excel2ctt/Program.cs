using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

/*! \mainpage 
        Công cụ nhập liệu từ file text lên trong CTT của HUST
*/
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
        /// <summary> Mã số sinh viên </summary>        
        public int MSSV;
        public string Name;
        public string Email;
        public string NgaySinh;        
        public double Grade;
        public string field1;
        public string field2;
        public string field3;

        public StudentInformation()
        {
            list.Add(this);
        }

        public StudentInformation(int STT, int MSSV, string Name, double Grade, string Email = null, string NgaySinh= null, string filed1 = null, string filed2 = null, string filed3 = null)
        {
            StudentInformation NewItem = new StudentInformation();
            NewItem.STT = STT;
            NewItem.Grade = Grade;
            NewItem.MSSV = MSSV;
            NewItem.Name = Name;
            NewItem.NgaySinh=NgaySinh;
            NewItem.Email = Email;            
            NewItem.field1 = field1;
            NewItem.field2 = field2;
            NewItem.field3 = field3;
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
        ///     Cấu trúc chuỗi thông tin đầu ra.
        /// </summary>
        /// <example>
        ///     email{tab}diem 
        /// </example>
        public static string OutFields;        

        /// <summary>
        ///      Nhập điểm
        /// </summary>
        /// <param name="values"> mảng chứa thông tin điểm, tuân theo đúng trình tự như <see cref="FieldNames"/> </param>
        /// 
        public StudentInformationBuilder(string[] values)
        {
            int i;
            for (i=0; i < values.Length; i++)
            {
                try
                {
                    if (FieldNames[i] == "diem")
                    {
                        Grade = Convert.ToDouble(values[i]);
                    }
                    else if (FieldNames[i] == "stt")
                    {
                        STT = Convert.ToInt32(values[i]);
                    }
                    else if (FieldNames[i] == "ten")
                    {
                        Name = values[i].Trim();
                    }
                    else if (FieldNames[i] == "mssv")
                    {
                        MSSV = Convert.ToInt32(values[i]);
                    }
                    else if (FieldNames[i] == "ngaysinh")
                    {
                        NgaySinh = values[i].Trim();
                    }
                    else if (FieldNames[i] == "email")
                    {
                        Email = values[i].Trim();
                    }
                    else if (FieldNames[i] == "field1")
                    {
                        field1 = values[i].Trim();
                    }
                    else if (FieldNames[i] == "field2")
                    {
                        field2 = values[i].Trim();
                    }
                    else if (FieldNames[i] == "field3")
                    {
                        field3 = values[i].Trim();
                    }
                }
                catch
                {
                    Console.WriteLine("Error: " + "column " + FieldNames[i] + " is invalid. " + values[i] );
                }
            }
        }

        /// <summary>
        ///      Chuỗi đầu ra đề nhập liệu
        /// </summary>
        /// <param name="values"> mảng chứa thông tin điểm, tuân theo đúng trình tự như <see cref="FieldNames"/> </param>
        static public string OutputString(StudentInformation si)
        {
            string myOutput = StudentInformationBuilder.OutFields;

            myOutput = myOutput.Replace("mssv", si.MSSV.ToString());
            myOutput = myOutput.Replace("email", si.Email);
            myOutput = myOutput.Replace("stt", si.STT.ToString());
            myOutput = myOutput.Replace("ten", si.Name);
            myOutput = myOutput.Replace("diem", si.Grade.ToString());
            myOutput = myOutput.Replace("ngaysinh", si.NgaySinh);
            myOutput = myOutput.Replace("field1", si.field1);
            myOutput = myOutput.Replace("field2", si.field2);
            myOutput = myOutput.Replace("field3", si.field3);
            return myOutput;
        }
    }
           

    class Program
    {


        static void Main(string[] args)
        {

           string[] fields;   //để phân tách các cột trong dòng thông tin
            char[] splitter; // các kí tự phân tách

            // Phân tích chuỗi cú pháp để biết ý nghĩa của các cột thông tin. Ví dụ   <name>;<grade>
            StudentInformationBuilder.FieldNames = Properties.Settings.Default.INPUTFORMAT.Split(new char[]{ ' ', ',', ';', '\t' });

            StudentInformationBuilder.OutFields = Properties.Settings.Default.OUTPUTFORMAT;

            /// Tạo đối tượng cần nhập điểm
            Victim site = new Victim(Properties.Settings.Default.AUTOACTIVATE);

            splitter = Properties.Settings.Default.SEPERATOR.Replace("\\t","\t").ToCharArray();

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
                        fields = line.Split(splitter);
                        // đưa vào cấu trúc để bóc tách và tự bổ sung vào danh sách
                        new StudentInformationBuilder(fields);
                    }

                }

                /// Tự động Activate site nạn nhân để tiền kiểm tra.
                site.Activate();

                /// Quay trở lại tiến trình hiện thời để nhìn rõ thông báo
                Victim.SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);

                /// Hiện thông báo confirm với người dùng
                DialogResult res = MessageBox.Show("Có " + lines.Length + " dòng nhập liệu. \nSau khi bấm OK, bạn có " + Properties.Settings.Default.WAITFORSWITCHINGAPP / 1000 + " giây để chuyển sang website ctt-sis và đặt con trỏ chuột vào ô điểm đầu tiên cần nhập. Để bỏ qua và kết thúc, bấm Cancel.\n" + lines[0] ?? "" + "\n" + lines[1] ?? "" + "\n" + lines[2] ?? "", "Chuẩn bị", MessageBoxButtons.OKCancel);
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
                else
                {
                    break;
                }
            } while (true);
        }
    }
}
