//#define EDX

using System;
using System.Collections.Generic;
using System.Threading;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel2ctt
{

    /// <summary>
    ///    Đối tượng chịu sự điều khiển. Ví dụ trang ctt-daotao, ictsv...
    /// </summary>
    class Victim
    {
        /// <summary>
        ///     
        /// </summary>
        /// <param name="hWnd">handle của cửa sổ. Lưu ý: phải dùng MainWindowHandle mới là của tiến trình cửa sổ, không dùng Handle vì là của tiến trình cha</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        ///     Tiến trình chạy website đang cần kiểm soát. Ví dụ chrome.exe hoặc là edge.exe
        /// </summary>
        private Process VictimProcess = null;

        public static List<StudentInformation> StudentList = null;

        public void AutoTyping()
        {
            StudentInformation si;

            /// Tự động activate cửa số nếu đã khai báo tiêu đề của cửa sổ
            Activate();

            /// Đợi trong khoảng thời gian đã ghi trong file cấu hình
            Thread.Sleep(Properties.Settings.Default.WAITFORSWITCHINGAPP);

            /// Nhập liệu
            for (int i = 0; i < StudentList.Count; i++)
            {
                si = StudentList[i];
#if EDX
                SendKeys.SendWait(si.Name.ToString());
                SendKeys.SendWait(Properties.Settings.Default.GOTONEXTRECORD);
                SendKeys.SendWait(si.field1.ToString());
                SendKeys.SendWait(Properties.Settings.Default.GOTONEXTRECORD);
                SendKeys.SendWait(si.field2.ToString());
                SendKeys.SendWait(Properties.Settings.Default.GOTONEXTRECORD);
                SendKeys.SendWait(si.field3.ToString());

#else
                SendKeys.SendWait(StudentInformationBuilder.OutputString(si));
#endif
                // Đợi một tí
                Thread.Sleep(Properties.Settings.Default.WAITFORNEXTRECORD);
                // Chuyển tới bản ghi kế tiếp
                SendKeys.SendWait(Properties.Settings.Default.GOTONEXTRECORD);

#if EDX
                SendKeys.SendWait("{Enter}");
                if (MessageBox.Show("Tiếp tục nhé","Hỏi đáp", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                {
                    break;
                }
#endif
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="Title">Một phần của tiêu đề cửa sổ nhập liệu. = null nếu bỏ qua. <> null thì chương trình sẽ tự động activate cửa số này trước khi nhập liệu.</param>
        public Victim(string Title = null)
        {
            if (Title == null || Title == "")
            {
                return;
            }

            /// Lấy ra danh sách các tiến trình đang chạy
            Process[] processlist = Process.GetProcesses();

            /// Tìm xem có tiến trình nào có tên tương tự không
            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {

                    if (process.MainWindowTitle.Contains(Title))
                    {
                        VictimProcess = process;
                        Console.WriteLine("Process:   {0}", process.ProcessName);
                        Console.WriteLine("    ID   : {0}", process.Id);
                        Console.WriteLine("    Title: {0} \n", process.MainWindowTitle);
                        break;
                    }
                }
            }
        }

        /// <summary>
        ///     Activate cửa số của tiến trình đang chạy
        /// </summary>
        public void Activate()
        {
            if (VictimProcess != null)
            {
                SetForegroundWindow(VictimProcess.MainWindowHandle);
            }
        }
    }
}
