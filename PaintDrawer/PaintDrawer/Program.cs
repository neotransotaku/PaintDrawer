using System;
using PaintDrawer.Letters;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using WindowScrape.Types;
using System.Windows.Input;
using PaintDrawer.Actions;
using System.Collections.Generic;
using PaintDrawer.GMail;

namespace PaintDrawer
{
    class Program
    {
        public static Stopwatch watch;
        public static double Time { get { return watch.Elapsed.TotalSeconds; } }
        public static double LastDraw = 0;
        public static CharFont font;
        public static Queue<IAction> queue;
        private static string Basepath = "D:\\Cygwin\\home\\mle\\2020_Q4_Fall\\PCC149\\DailyReport";

        // Windows comlains if this isnt decorated with a STAThread
        [STAThread]
        static void Main(string[] args)
        {
            if ((int)DateTime.Now.DayOfWeek == 3)
            {
                Console.ForegroundColor = Colors.Success;
                Console.WriteLine("it is wednesday, my dudes\n");
            }

            Console.ForegroundColor = Colors.Success;
            Console.WriteLine("Welcome to Paint Drawer V0.1!");
            //The drawing and stuff is in another thread
            Thread t = new Thread(Soreni);
            t.Start();

            // The most cancer way to kill the program; literally kill it from another thread.
            // Dont judge the code someone made way past midnight
            while (true)
            {
                // Just press ASD at the same time and boom, you just fucked everything up :D
                if (Keyboard.IsKeyDown(Key.A) && Keyboard.IsKeyDown(Key.S) && Keyboard.IsKeyDown(Key.D))
                {
                    Console.ForegroundColor = Colors.Error;
                    Console.WriteLine("Keep holding ASD for 3 more seconds and the program dies!");

                    Stopwatch watch = Stopwatch.StartNew();
                    while (Keyboard.IsKeyDown(Key.A) && Keyboard.IsKeyDown(Key.S) && Keyboard.IsKeyDown(Key.D))
                    {
                        if (watch.ElapsedMilliseconds > 2000)
                        {
                            t.Abort();
                            Console.ForegroundColor = Colors.Error;
                            Console.WriteLine("[Main] Program has been killed. Easy GG, amiright?");
                            Input.MouseUp();
                            Console.ReadLine();
                            return;
                        }
                        Thread.Sleep(1);
                    }
                    Console.ForegroundColor = Colors.Message;
                    Console.WriteLine("Program abort canceled.");
                }
                Thread.Sleep(10);
            }
        }

        static void Soreni()
        {
            DateTime now = DateTime.Now;
            string today = String.Format("{0}-{1:00}-{2:00}", now.Year, now.Month, now.Day);
            if(!Directory.Exists(Basepath + "\\" + today))
            {
                Basepath = Basepath + "\\" + today;
                Directory.CreateDirectory(Basepath + "\\" + today);
            } else
            {
                Basepath = Basepath + "\\" + today;
            }

            string timestamp = String.Format("{0}{1:00}{2:00}_{3:00}{4:00}", now.Year, now.Month, now.Day, now.Hour, now.Minute);
            AxiUmPuppet axiUm = AxiUmPuppet.GetInstance();

            List<Tuple<string, string>> ranges = new List<Tuple<string, string>>();
            ranges.Add(Tuple.Create("S2100", "S2299"));
            ranges.Add(Tuple.Create("I2100", "I2299"));
            axiUm.RunAppointmentReport(Basepath + "\\appointment-" + timestamp, "07/01/2020", "06/30/2021", ranges);
            axiUm.RunChairReport(Basepath + "\\chair-" + timestamp, "07/01/2020", "06/30/2021", ranges);
            axiUm.RunChargeReport(Basepath + "\\charge-" + timestamp, "07/01/2019", String.Format("{1:00}/{2:00}/{0}", now.Year, now.Month, now.Day), ranges);
            axiUm.RunFeedbackReport(Basepath + "\\feedback-"+timestamp, "07/01/2019", String.Format("{1:00}/{2:00}/{0}", now.Year, now.Month, now.Day), ranges);
            foreach(string practice in new string[] { "5", "6", "9" })
            {
                axiUm.RunInProcessReport(Basepath + "\\inprocess["+practice+"]-" + timestamp, "07/01/2019", String.Format("{1:00}/{2:00}/{0}", now.Year, now.Month, now.Day), ranges, practice);
            }
            Thread.Sleep(500);

            Console.ForegroundColor = Colors.Message;
            Console.WriteLine("Task is done");

            ShowSelf();
            Input.RegisterKeyDown(Input.KEY_A);
            Input.RegisterKeyDown(Input.KEY_S);
            Input.RegisterKeyDown(Input.KEY_D);

            return;
        }

        static void ShowSelf()
        {
            List<HwndObject> list = HwndObject.GetWindows();
            foreach (HwndObject o in list)
            {
                if (o.Size.Width == 0 || o.Size.Height == 0 || o.Text.Length == 0)
                    continue;

                if (o.ClassName.Equals("ConsoleWindowClass") || o.Title.Contains("PaintDrawer.exe"))
                {
                    WindowScrape.Static.HwndInterface.ShowWindow(o.Hwnd, 3);
                }
            }
        }

        // This method processes the main program. Main() ends up keeping track of whether it should kill
        // the program while this guy does all the actual hard work of drawing, checking mails, etc.
        static void ProcessStuff()
        {
            Stuff.Init();

            Account.Init();

            // Loads the font into memory from the .map files
            font = new CharFont("CharFiles/MizConsole/");
            Actions.Actions.Init(font);

            Thread.Sleep(1000);

            HwndObject obj;
            if (GetPaintHwnd(out obj))
            {
                Console.ForegroundColor = Colors.Success;
                Console.WriteLine("[Program] Paint Hwnd found! using: " + obj.Text);
                obj.Activate();
                WindowScrape.Static.HwndInterface.ShowWindow(obj.Hwnd, 3);
            }
            else
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("[Program] Paint Hwnd not found.");
                Input.OpenPaint();
                if (!ForceGetPaintHwnd(out obj))
                {
                    Console.ForegroundColor = Colors.Error;
                    Console.WriteLine("[Program] ERROR: Can't find Paint Hwnd. Process aborted.");
                    return;
                }
            }

            Thread.Sleep(1000);

            if (obj.Location.X > 10 && obj.Location.Y > 10 && obj.Size.Width < Stuff.ScreenWidth - 50 && obj.Size.Height < Stuff.ScreenHeight - 50)
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("Paint window not maximized, maximizing...");
                Input.MoveTo(new Point(obj.Location.X + obj.Size.Width - 69, obj.Location.Y + 3));
                Input.RegisterClick();
                Thread.Sleep(100);
            }

            watch = Stopwatch.StartNew();
            queue = new Queue<IAction>(64);
            Console.ForegroundColor = Colors.Success;
            Console.WriteLine("Stopwatch started. Entering main loop...");

            Input.PaintSelectBrush();
            new SimpleWrite(font, "Welcome to PaintDrawer! What shall I draw for you, kind sir?", 70).Act();

            System.Text.StringBuilder build = new System.Text.StringBuilder(256);
            for (int i = 32; i < CharFont.MaxCharNumericValue; i++)
                if (font.DoesCharExist((char)i))
                    build.Append((char)i);
            new SimpleWrite(font, build.ToString(), 70).Act();

            Actions.Actions.CreateSimpleWrite(font, "Este es un mensaje especial! Estas durmiendo en un estado de coma, metimos esto en tu cabeza pero no sabemos a donde va a terminar, no tenemos idea que hacer con vos, por favor despertate! no porque te extrañemos pero yo que se...").Act();
            LastDraw = -60;

            while (true)
            {
                Account.AddNewToQueue(queue);

                if (queue.Count == 0 && Time - LastDraw > 5 * 60)
                {
                    // More than X minutes passed since the last draw! Let's add something random then...
                    queue.Enqueue(Actions.Actions.RandomAction);
                }

                if (queue.Count != 0 && Time - LastDraw > 1 * 60)
                {
                    // Give all texts at least a minute before erasing 'em... and drawing something else
                    Input.PaintClearImage();
                    Input.PaintSelectBrush();
                    LastDraw = Time;
                    queue.Dequeue().Act();
                }
                else
                    Thread.Sleep(2000);

            }
        }

        /// <summary>
        /// Tries 5 times to get a HwndObject corresponding to Paint's window. Returns true if one was found, false (and null) otherwise.
        /// </summary>
        /// <param name="obj">The HwndObject found</param>
        static bool ForceGetPaintHwnd(out HwndObject obj)
        {
            const int MaxTries = 5;

            List<HwndObject> list = HwndObject.GetWindows();
            int i = 0;
            while (true)
            {
                if (GetPaintHwnd(out obj))
                {
                    Console.ForegroundColor = Colors.Success;
                    Console.WriteLine("[Program] Found (hopefully) paint window: " + obj.Text);
                    return true;
                }

                if (++i >= MaxTries)
                {
                    Console.WriteLine(String.Concat("[Program] (", i, ')', " No Paint window found. Not retrying #yolo"));
                    obj = null;
                    return false;
                }
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine(String.Concat("[Program] (", i, ")", "No Paint window found. Retrying in 1s..."));
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// Tries once to get a HwndObject with the string "Paint". Returns whether any was found
        /// </summary>
        /// <param name="obj">The Hwnd object found</param>
        /// <returns></returns>
        static bool GetPaintHwnd(out HwndObject obj)
        {
            List<HwndObject> list = HwndObject.GetWindows();
            foreach (HwndObject o in list)
            {
                if (o.Text.Contains("Paint") && !o.Text.ToLower().Contains("drawer"))
                {
                    obj = o;
                    return true;
                }
            }
            obj = null;
            return false;
        }
    }
}
