using bot_minsa.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bot_minsa
{
    class Program
    {
        static void Main(string[] args)
        {
            cls_Process oProcess = new cls_Process();

            //while (true)
            //{


            //    //Console.Clear();

            //    DateTime now = System.DateTime.Now;
            //    var hour = now.Hour;
            //    var minute = now.Minute;
            //    var current_time = hour + minute / 60;
            //    Console.WriteLine(now);
            //    Console.WriteLine("---");
            //    Console.WriteLine(current_time);

            //    if (current_time <= 5)
            //    {
            //        Console.WriteLine("wait");
            //        System.Threading.Thread.Sleep(1800000);
            //    }
            //    else
            //    {
            //        Console.WriteLine("run");
            
            oProcess.start_Process();
            System.Environment.Exit(0);
            return;
            //        System.Threading.Thread.Sleep();


            //    }
            //    Console.WriteLine(now);


            //}


        }
    }
}
