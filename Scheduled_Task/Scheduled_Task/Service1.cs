using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace Scheduled_Task
{
    public partial class Service1 : ServiceBase
    {
        System.Timers.Timer timer1;  //计时器

        int StartHour1 = int.Parse(ConfigurationManager.AppSettings["Scheduled_Task_ExecuteTime_1"].Split(':')[0]);
        int StartMinute1 = int.Parse(ConfigurationManager.AppSettings["Scheduled_Task_ExecuteTime_1"].Split(':')[1]);
        int StartSecond1 = int.Parse(ConfigurationManager.AppSettings["Scheduled_Task_ExecuteTime_1"].Split(':')[2]);

        string LogPath = ConfigurationManager.AppSettings["LogPath"];
        string ExportPath = ConfigurationManager.AppSettings["ExportPath"];

        string SQLstr = ConfigurationManager.AppSettings["SQLstr"];

        public Service1()
        {
            InitializeComponent();
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            //记录日志
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "OnStart.txt", true))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Start.");
            }
            timer1 = new System.Timers.Timer();
            timer1.Interval = 1000;  //设置计时器事件间隔执行时间
            timer1.Elapsed += new System.Timers.ElapsedEventHandler(TMStart1_Elapsed);
            timer1.Enabled = true;
        }

        protected override void OnStop()
        {
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "OnStop.txt", true))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Stop.");
            }
            this.timer1.Enabled = false;
        }

        protected override void OnPause()
        {
            //服务暂停执行代码
            base.OnPause();
        }
        protected override void OnContinue()
        {
            //服务恢复执行代码
            base.OnContinue();
        }
        protected override void OnShutdown()
        {
            //系统即将关闭执行代码
            base.OnShutdown();
        }

        private void TMStart1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
#if DEBUG
#else
            int Hour = e.SignalTime.Hour;
            int Minute = e.SignalTime.Minute;
            int Second = e.SignalTime.Second;

            if ((Hour == StartHour1 && Minute == StartMinute1 && Second == StartSecond1))
            {
#endif
                System.Timers.Timer obtain_time1 = (System.Timers.Timer)timer1;
                obtain_time1.Enabled = false;  //运行前停止执行Elapsed事件

                //记录日志
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "TaskLog.txt", true))
                {
                    sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Task Start.");
                }

                //需要定时运行的部分
                TaskPart();

                obtain_time1.Enabled = true;  ////运行完毕后开始执行Elapsed事件
#if DEBUG
#else
            }
#endif
        }

        /// <summary>
        /// 需要定时运行的部分
        /// </summary>
        private void TaskPart()
        {
            string sError = string.Empty;
            DataTable dt = SQLServerHelper.GetDataTable(out sError, SQLstr);

            if (dt.Rows.Count > 0)
            {
                bool ExportStatus = CSVHelper.dt2csvClient(dt, ExportPath + "[Data]" + DateTime.Now.ToString().Replace("/","").Replace(":","").Replace(" ","")+ ".csv");
                if (ExportStatus == true)
                {
                    //记录日志
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "ExportStatusLog.txt", true))
                    {
                        sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Export Successfully.");
                    }
                }
                else
                {
                    //记录日志
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "ExportStatusLog.txt", true))
                    {
                        sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Export Failed.");
                    }
                }

            }
            else
            {
                //记录日志
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(LogPath + "ExportStatusLog.txt", true))
                {
                    sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "No Data Selected.");
                }
            }
        }
    }
}
