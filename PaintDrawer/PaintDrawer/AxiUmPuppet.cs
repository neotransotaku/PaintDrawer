using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using WindowScrape.Types;

namespace PaintDrawer
{
    class AxiUmPuppet
    {
        private HwndObject axiUm;

        private const int DEFAULT_RETRIES = 2;

        public AxiUmPuppet(HwndObject axiUm)
        {
            Stuff.Init();
            Thread.Sleep(1000);

            this.axiUm = axiUm;
            axiUm.Activate();
            WindowScrape.Static.HwndInterface.ShowWindow(axiUm.Hwnd, 3);
            Thread.Sleep(1000);

            // Maximize AxiUm if it already isn't
            axiUm.Maximize();

            // Close any open reports to ensure axiUm is in its default resting state
            ClearOpenReports();
        }

        public bool CloseInfoManager()
        {
            foreach (HwndObject o in getMainContainer().GetChildren())
            {
                Console.ForegroundColor = Colors.Message;
                if (o.Title.StartsWith("Info Manager"))
                {
                    o.CloseWindow();
                    return true;
                }
            }

            return false;
        }

        public HwndObject GetInfoManager() // TODO: make this a retryable action
        {
            foreach (HwndObject o in getMainContainer().GetChildren())
            {
                Console.ForegroundColor = Colors.Message;
                if (o.Title.StartsWith("Info Manager"))
                {
                    Stuff.WriteConsoleMessage("Info Manager already opened");
                    return o;
                }
            }

            Stuff.WriteConsoleMessage("Info Manager not opened...");
            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_A);
            Thread.Sleep(1000);
            Input.PressKey(Input.KEY_I);
            Thread.Sleep(1000);
            HwndObject infoManager = GetInfoManager();
            if(infoManager == null)
            {
                Stuff.WriteConsoleError("Unable to open info manager");
            }
            return infoManager;
        }

        private HwndObject getMainContainer()
        {
            List<HwndObject> list = axiUm.GetChildren();
            if (list.Count != 3)
            {
                Console.ForegroundColor = Colors.Error;
                foreach (HwndObject o in list)
                {
                    Console.WriteLine(Stuff.GetHwndInfoString(o));
                }
                throw new Exception("Axium changed its top level children count");
            }
            else
            {
                var mainContainer = list[1];
                if (!mainContainer.ClassName.Equals("MDIClient"))
                {
                    Console.ForegroundColor = Colors.Message;
                    foreach (HwndObject o in list)
                    {
                        Console.WriteLine(Stuff.GetHwndInfoString(o));
                    }
                    Console.WriteLine("It is unclear if this window is the one you want");
                }
                return mainContainer;
            }
        }

        private void ClearOpenReports()
        {
            List<HwndObject> list = HwndObject.GetWindows().FindAll(o => o.Title.Equals("Enter Parameter Values"));
            foreach (HwndObject o in list)
            {
                o.CloseWindow();
                Thread.Sleep(1000);
                Input.PressKey(Input.KEY_SPACE);
                axiUm.Activate();
            }
        }

        public void RunAppointmentReport(string path, string startDate, string endDate, List<Tuple<string, string>> studentRanges, int numRetries = DEFAULT_RETRIES)
        {
            if (!EnoughRetries(numRetries))
                return;

            CloseInfoManager();
            if (GetInfoManager() == null)
            {
                Stuff.WriteConsoleError("Unable to open info manager...Giving Up");
                return;
            }

            Input.MoveTo(new Point(480, 95)); // Appointment Tab
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(500);

            Input.MoveTo(new Point(1060, 170)); // Pre-Defined Button
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(1000);

            Input.MoveTo(new Point(845, 500)); // Michael's Appointment List
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(1000);

            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_S); // Select the report

            Stuff.WriteConsoleMessage("Setting date range of " + startDate + " to " + endDate);
            SetDateFieldParameters(new Point(335, 350), startDate, endDate);
            Thread.Sleep(500);

            Stuff.WriteConsoleMessage("Setting student ranges");
            studentRanges.Add(Tuple.Create("D106", "D106")); //NPE
            studentRanges.Add(Tuple.Create("D099", "D099")); //ER
            SetStudentRangeParameters(new Point(560, 350), studentRanges);
            Thread.Sleep(500);

            if (!new RetryableAction(() => ExecuteQuery(), "Executing Query").Execute())
            {
                Stuff.WriteConsoleError("Unable to execute query to gather appointments");
                return;
            }
            Thread.Sleep(5000);

            GetInfoManager().Activate();
            if (!new RetryableAction(() => ExportQuery(), "Exporting Query Results").Execute())
            {
                Stuff.WriteConsoleError("Unable to export query results");
                return;
            }
            Thread.Sleep(5000);

            if (!new RetryableAction(() => SaveExcel(path), "Saving Generated Excel File").Execute())
            {
                Stuff.WriteConsoleError("Unable to save excel file");
                return;
            }
            Thread.Sleep(5000);
        }

        private void ExecuteQuery()
        {
            HwndObject startWindow = HwndObject.GetForegroundWindow();

            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_S);
            Stopwatch watch = Stopwatch.StartNew();
            Thread.Sleep(500);

            HwndObject fetching = null;
            foreach (HwndObject o in HwndObject.GetWindows())
            {
                List<HwndObject> list = o.GetChildren();
                if (list.Count != 3)
                    continue;

                if (list[1].Title.Equals("Please wait...") || list[1].Title.Equals("Fetching Data..."))
                {
                    Stuff.WriteConsoleMessage("Parent: " + Stuff.GetHwndInfoString(o.GetParent()));
                    fetching = list[0];
                    break;
                }
            }

            while (HwndObject.GetForegroundWindow() != startWindow)
            {
                Console.Write("Waiting... (" + watch.ElapsedMilliseconds + ")");
                if(fetching != null)
                {
                    Console.WriteLine(fetching.Text);
                } else
                {
                    Console.WriteLine();
                }
                Thread.Sleep(5000);
            }
            watch.Stop();

            Stuff.WriteConsoleSuccess("Report Generated");
        }

        private void ExportQuery()
        {
            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_A);
            Thread.Sleep(5000);

            HwndObject startWindow = HwndObject.GetForegroundWindow();

            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_E);
            Thread.Sleep(1000);

            Input.MoveTo(new Point(870, 405));
            Input.RegisterClick();
            Thread.Sleep(500);
            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_O);

            Stopwatch watch = Stopwatch.StartNew();
            Thread.Sleep(10000);

            HwndObject progress = null;
            foreach (HwndObject o in HwndObject.GetWindows())
            {
                List<HwndObject> list = o.GetChildren();
                if (list.Count <= 1)
                    continue;

                if (list[1].Title.Equals("Processing"))
                {
                    Stuff.WriteConsoleMessage(Stuff.GetHwndInfoString(list[0]));
                    progress = list[0];
                    break;
                }
            }

            int excelWindowCount = HwndObject.GetWindows().FindAll(e => e.Title.Contains("Excel")).Count;
            while (true)
            {
                HwndObject window = HwndObject.GetForegroundWindow();
                if (window.Title.Contains("Excel") && window.Title.StartsWith("Book"))
                {
                    Stuff.WriteConsoleMessage("Found opened window");
                    break;
                }
                else if(HwndObject.GetWindows().FindAll(e => e.Title.Contains("Excel")).Count == (excelWindowCount + 1))
                {
                    Stuff.WriteConsoleMessage("Detected change in excel window counts");
                    break;
                }

                Console.Write("Waiting... (" + watch.ElapsedMilliseconds + ")");
                if (progress != null)
                {
                    Console.WriteLine(" " + progress.Text);
                } else
                {
                    Console.WriteLine();
                }
                Thread.Sleep(5000);
            }
            watch.Stop();

            Stuff.WriteConsoleSuccess("Data exported");
        }

        public void SaveExcel(string path)
        {
            HwndObject window = HwndObject.GetForegroundWindow();
            if (!window.Title.Contains("Excel") && !window.Title.StartsWith("Book"))
            {
                foreach (HwndObject o in HwndObject.GetWindows().FindAll(e => e.Title.Contains("Excel")))
                {
                    if (o.Title.StartsWith("Book"))
                    {
                        if (window == null)
                        {
                            window = o;
                        }
                        else if (Int32.TryParse(o.Title.Substring(4), out int number))
                        {
                            if (number > Int32.Parse(window.Title.Substring(4)))
                                window = o;
                        }
                    }
                }
            }

            if (window == null)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("Unable to export report from Axium");
                return;
            }
            else
            {
                Console.ForegroundColor = Colors.Success;
                Console.WriteLine("Excel file opened: " + window.Title);
            }

            window.Activate();
            Thread.Sleep(2000);
            window.Maximize();
            Thread.Sleep(2000);
            Input.PressKeyCombo(Input.KEY_CONTROL, Input.KEY_S);
            Thread.Sleep(2000);

            Input.MoveTo(new Point(800, 635));
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(500);

            Input.MoveTo(new Point(570, 195));
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(500);

            HwndObject saveDialog = HwndObject.GetWindows().Find(e => e.Title.Contains("Save As"));
            Stuff.WriteConsoleMessage("Save dialog found!");
            saveDialog.Activate();
            Thread.Sleep(500);
            Input.KeyboardWrite(path, 100);
            Thread.Sleep(100);
            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_S);
            Thread.Sleep(2000);

            if (!File.Exists(path + ".xlsx"))
            {
                Stuff.WriteConsoleError("File was not saved...trying again");
            }
            else
            {
                window.CloseWindow();
                Stuff.WriteConsoleSuccess("File was saved to " + path);
            }
        }

        private bool OpenColumnCondition(Point p, int numRetries = DEFAULT_RETRIES)
        {
            if (!EnoughRetries(numRetries))
                return false;

            Input.MoveTo(p);
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(2000);

            HwndObject conditions = null;
            foreach (HwndObject o in HwndObject.GetWindows())
            {
                if (o.Title.Equals("Column Conditions"))
                {
                    conditions = o;
                    break;
                }
            }
            if (conditions == null)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("Unable to find Column Conditions box");
                Console.WriteLine("Not enough time to find these windows. Trying once more...");
                return OpenColumnCondition(p, numRetries - 1);
            }
            else
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("Column Conditions box found");
                conditions.Activate();
            }

            return true;
        }

        private void SetStudentRangeParameters(Point p, List<Tuple<string, string>> studentRanges, int numRetries = DEFAULT_RETRIES)
        {
            if (!EnoughRetries(numRetries) || !OpenColumnCondition(p))
            {
                Stuff.WriteConsoleError("Unable to set student range parameters");
                return;
            }

            Input.PressKey(Input.KEY_B);
            Input.PressKey(Input.KEY_TAB);

            int count = 0;
            foreach (Tuple<string, string> range in studentRanges)
            {
                Input.KeyboardWrite(range.Item1, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.KeyboardWrite(range.Item2, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_ENTER);

                if((count + 1) != studentRanges.Count)
                {
                    Input.RegisterKeyDown(Input.KEY_SHIFT);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.RegisterKeyUp(Input.KEY_SHIFT);
                    count++;
                }
            }
            Input.RegisterKeyDown(Input.KEY_ALT);
            Input.PressKey(Input.KEY_C);
            Input.RegisterKeyUp(Input.KEY_ALT);
        }

        private void SetDateFieldParameters(Point p, string startDate, string endDate, int numRetries = DEFAULT_RETRIES)
        {
            if (!EnoughRetries(numRetries) || !OpenColumnCondition(p))
            {
                Stuff.WriteConsoleError("Unable to set date range parameters");
                return;
            }

            Input.PressKey(Input.KEY_B); // Set to "Between"
            Input.PressKey(Input.KEY_TAB);
            Input.KeyboardWrite(startDate, 100);
            Input.PressKey(Input.KEY_TAB);
            Input.PressKey(Input.KEY_TAB);
            Input.KeyboardWrite(endDate, 100);
            Input.PressKey(Input.KEY_TAB);
            Input.PressKey(Input.KEY_TAB);
            Input.PressKey(Input.KEY_ENTER);
            Input.PressKeyCombo(Input.KEY_ALT, Input.KEY_C);
        }

        public void RunChairReport(string path, string startDate, string endDate, List<Tuple<string, string>> studentRanges, int numRetries = DEFAULT_RETRIES)
        {
            if (!EnoughRetries(numRetries))
                return;

            CloseInfoManager();
            if (GetInfoManager() == null)
            {
                Stuff.WriteConsoleError("Unable to open info manager...Giving Up");
                return;
            }

            Input.MoveTo(new Point(800, 95)); // Custom Lists
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(500);

            Input.MoveTo(new Point(300, 220)); // Category
            Thread.Sleep(500);
            Input.RegisterClick();

            Input.MoveTo(new Point(485, 350));
            Thread.Sleep(500);
            Input.RegisterClick();
            Thread.Sleep(500);

            Stuff.WriteConsoleMessage("Setting student ranges");
            SetStudentRangeParameters(new Point(485, 350), studentRanges);
            Thread.Sleep(500);

            Stuff.WriteConsoleMessage("Setting date range of " + startDate + " to " + endDate);
            SetDateFieldParameters(new Point(710, 350), startDate, endDate);
            Thread.Sleep(500);

            if (!new RetryableAction(() => ExecuteQuery(), "Executing Query").Execute())
            {
                Stuff.WriteConsoleError("Unable to execute query to gather appointments");
                return;
            }
            Thread.Sleep(5000);

            GetInfoManager().Activate();
            if (!new RetryableAction(() => ExportQuery(), "Exporting Query Results").Execute())
            {
                Stuff.WriteConsoleError("Unable to export query results");
                return;
            }
            Thread.Sleep(5000);

            if (!new RetryableAction(() => SaveExcel(path), "Saving Generated Excel File").Execute())
            {
                Stuff.WriteConsoleError("Unable to save excel file");
                return;
            }
            Thread.Sleep(5000);
        }

        public void RunChargeReport(string path, string startDate, string endDate, List<Tuple<string, string>> studentRanges, int numRetries = DEFAULT_RETRIES)
        {
            Action findReport = () =>
            {
                Input.MoveTo(new Point(530, 205));
                Thread.Sleep(500);
                Input.RegisterClick();
                Thread.Sleep(500);
                Input.PressKey(Input.KEY_END);
                Thread.Sleep(500);
                Input.RegisterClick();
                Thread.Sleep(500);
            };

            Action reportParameters = () =>
            {
                // Clinic Select
                Input.MoveTo(new Point(645, 345));
                Input.RegisterClick();
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_ENTER);    // Load Parnassus
                Input.PressKey(Input.KEY_ENTER);    // Load Buchanan
                Input.RegisterKeyDown(Input.KEY_SHIFT);
                Input.PressKey(Input.KEY_TAB);
                Input.RegisterKeyUp(Input.KEY_SHIFT);
                Input.PressKey(Input.KEY_DOWN);
                Thread.Sleep(500);
                Input.PressKey(Input.KEY_TAB); 
                Input.PressKey(Input.KEY_ENTER);    // Load Perio

                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);

                // Date Select
                Input.KeyboardWrite(startDate, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.KeyboardWrite(endDate, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);

                // Select Producers
                int count = 0;
                foreach (Tuple<string, string> range in studentRanges)
                {
                    Input.KeyboardWrite(range.Item1, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.KeyboardWrite(range.Item2, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_ENTER);

                    if ((count + 1) != studentRanges.Count)
                    {
                        Input.RegisterKeyDown(Input.KEY_SHIFT);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.RegisterKeyUp(Input.KEY_SHIFT);
                        count++;
                    }
                }
                Thread.Sleep(100);

                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_ENTER);
            };

            RunReport("Prod Stat by Provider by Code", path, findReport, reportParameters);
            Thread.Sleep(5000);
        }

        public void RunFeedbackReport(string path, string startDate, string endDate, List<Tuple<string, string>> studentRanges, int numRetries = DEFAULT_RETRIES)
        {
            Action findReport = () =>
            {
                Input.MoveTo(new Point(500, 335));
                Input.RegisterClick();
                Input.PressKey(Input.KEY_HOME);
                Thread.Sleep(500);
                Input.RegisterClick();
                Thread.Sleep(500);
            };

            Action reportParameters = () =>
            {
                Input.MoveTo(new Point(700, 345));
                Input.RegisterClick();

                // Select Producers
                int count = 0;
                foreach (Tuple<string, string> range in studentRanges)
                {
                    Input.KeyboardWrite(range.Item1, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.KeyboardWrite(range.Item2, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_ENTER);

                    if ((count + 1) != studentRanges.Count)
                    {
                        Input.RegisterKeyDown(Input.KEY_SHIFT);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.RegisterKeyUp(Input.KEY_SHIFT);
                        count++;
                    }
                }
                Thread.Sleep(100);

                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);

                // Select Date
                Input.KeyboardWrite(startDate, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.KeyboardWrite(endDate, 100);
                Input.PressKey(Input.KEY_ENTER);
            };

            RunReport("Feedback Card", path, findReport, reportParameters);
            Thread.Sleep(5000);
        }

        public void RunInProcessReport(string path, string startDate, string endDate, List<Tuple<string, string>> studentRanges, string practiceNumber = "5", int numRetries = DEFAULT_RETRIES)
        {
            Action findReport = () =>
            {
                Input.MoveTo(new Point(520, 415));
                Input.RegisterClick();
                Input.PressKey(Input.KEY_HOME);
                Thread.Sleep(500);
                Input.RegisterClick();
                Thread.Sleep(500);
            };

            Action reportParameters = () =>
            {
                Input.MoveTo(new Point(690, 375));
                Input.RegisterClick();

                // Select Producers
                int count = 0;
                foreach (Tuple<string, string> range in studentRanges)
                {
                    Input.KeyboardWrite(range.Item1, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.KeyboardWrite(range.Item2, 100);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_TAB);
                    Input.PressKey(Input.KEY_ENTER);

                    if ((count + 1) != studentRanges.Count)
                    {
                        Input.RegisterKeyDown(Input.KEY_SHIFT);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.PressKey(Input.KEY_TAB);
                        Input.RegisterKeyUp(Input.KEY_SHIFT);
                        count++;
                    }
                }
                Thread.Sleep(100);

                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);

                // Select Practice
                Input.KeyboardWrite(practiceNumber, 100);
                Input.PressKey(Input.KEY_TAB);

                // Select Date
                Input.KeyboardWrite(startDate, 100);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.PressKey(Input.KEY_TAB);
                Input.KeyboardWrite(endDate, 100);
                Input.PressKey(Input.KEY_ENTER);
            };
            RunReport("In Process Tracking", path, findReport, reportParameters);
            Thread.Sleep(5000);
        }

        private void RunReport(string reportName, string path, Action findReport, Action reportParameters, int numRetries = 3)
        {
            if (!EnoughRetries(numRetries))
                return;

            CloseInfoManager();
            if (GetInfoManager() == null)
            {
                Stuff.WriteConsoleError("Unable to open info manager...Giving Up");
                return;
            }
            Thread.Sleep(1000);

            findReport.Invoke();
            Input.PressKey(Input.KEY_TAB);
            Thread.Sleep(500);
            Input.PressKey(Input.KEY_ENTER);
            Thread.Sleep(5000);

            Console.ForegroundColor = Colors.Message;
            Console.WriteLine("Loading "+reportName+", will save to " + path);

            HwndObject parameters = null;
            HwndObject report = null;
            if (!InitializeReport(out parameters, out report, reportName))
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("Trying again...");
                RunReport(reportName, path, findReport, reportParameters, numRetries - 1);
                return;
            }
            Thread.Sleep(5000);

            reportParameters.Invoke();

            Console.ForegroundColor = Colors.Message;
            Console.WriteLine("Requesting Report");

            Stopwatch watch = Stopwatch.StartNew();
            while (report.GetChildren()[0].GetChildren().Count == 0)
            {
                Console.WriteLine("Waiting... (" + watch.ElapsedMilliseconds + ")");
                Thread.Sleep(3000);
            }
            watch.Stop();

            Console.ForegroundColor = Colors.Success;
            Console.WriteLine("Report Generated");

            SaveReport(path);

            Console.ForegroundColor = Colors.Message;
            Console.WriteLine("Report Processed");

            report.CloseWindow();
        }

        private void SaveReport(string path, int numberRetries = 3)
        {
            if (!EnoughRetries(numberRetries))
                return;

            Console.ForegroundColor = Colors.Message;
            Console.WriteLine("Saving Report to " + path);

            Input.MoveTo(new Point(15, 40));
            Thread.Sleep(100);
            Input.RegisterClick();
            Thread.Sleep(2000);

            HwndObject export = null;
            foreach (HwndObject o in HwndObject.GetWindows())
            {
                if (o.Title.Equals("Export Report"))
                {
                    export = o;
                }
            }
            if (export == null)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("Unable to find export dialog box");
                Console.WriteLine("Not enough time to find these windows. Trying once more...");
                SaveReport(path, numberRetries-1);
                return;
            }
            else
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("Export dialog box found");
                export.Activate();
            }

            Input.KeyboardWrite(path, 100);
            Input.PressKey(Input.KEY_TAB);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_DOWN);
            Input.PressKey(Input.KEY_TAB);
            Input.PressKey(Input.KEY_ENTER);

            HwndObject confirm = null;
            Stopwatch watch = Stopwatch.StartNew();
            while(watch.ElapsedMilliseconds < 60000)
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("Saving... ("+watch.ElapsedMilliseconds+")");

                // Find confirmation window
                foreach (HwndObject o in HwndObject.GetWindows())
                {
                    if (o.Title.Equals("Export Report") && o.GetChildren().Count == 3)
                    {
                        confirm = o;
                        break;
                    }
                }

                if(confirm != null)
                {
                    Thread.Sleep(500);
                    confirm.Activate();
                    Input.PressKey(Input.KEY_ENTER);
                    break;
                } else
                {
                    Thread.Sleep(3000);
                }
            }
            watch.Stop();

            if (!File.Exists(path+".xlsx"))
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("File was not saved...trying again");
                SaveReport(path, numberRetries - 1);
            }
            else
            {
                Console.ForegroundColor = Colors.Success;
                Console.WriteLine("File was saved");
            }
        }

        private bool EnoughRetries(int numRetries)
        {
            if (numRetries == 0)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("No more retries available. Giving up.");
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool InitializeReport(out HwndObject parameters, out HwndObject report, string reportName)
        {
            parameters = null;
            report = null;

            foreach (HwndObject o in HwndObject.GetWindows())
            {
                if (o.Title.Equals(reportName))
                {
                    report = o;
                }
                else if (o.Title.Equals("Enter Parameter Values"))
                {
                    parameters = o;
                }
            }

            if (parameters == null || report == null)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("Not enough time to find these windows.");
                return false;
            }

            return true;
        }

        private static bool GetAxium(out HwndObject axiUm)
        {
            List<HwndObject> list = HwndObject.GetWindows();
            foreach (HwndObject o in list)
            {
                if (o.Size.Width == 0 || o.Size.Height == 0 || o.Text.Length == 0)
                    continue;

                if (o.Text.Contains("axiUm"))
                {
                    axiUm = o;
                    return true;
                }
            }
            axiUm = null;
            return false;

            /*
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
            */
        }

        public static AxiUmPuppet GetInstance(int numRetries = 1)
        {
            if (numRetries == 0)
            {
                Console.ForegroundColor = Colors.Error;
                Console.WriteLine("[Program] ERROR: Can't find AxiUm Hwnd. Process aborted.");
                return null;
            }


            Console.WriteLine("Loading axiUm");
            HwndObject axiUm;
            if (GetAxium(out axiUm))
            {
                Console.ForegroundColor = Colors.Success;
                Console.WriteLine("[Program] axiUm found! using: " + axiUm.Text);
                return new AxiUmPuppet(axiUm);
            }
            else
            {
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("[Program] axiUm Hwnd not found. Attempting to load manually.");
                Input.PressKey(Input.KEY_LWIN);
                Thread.Sleep(2000);
                Input.KeyboardWrite("axium", 100);
                Thread.Sleep(2000);
                Input.PressKey(Input.KEY_ENTER);
                Thread.Sleep(10000);
                Console.WriteLine("[Program] Logging in...");
                Input.KeyboardWrite("username", 100);
                Input.PressKey(Input.KEY_TAB);
                Input.KeyboardWrite("password", 100);
                Input.PressKey(Input.KEY_ENTER);
                Thread.Sleep(10000);
                /*
                Input.OpenPaint();
                if (!ForceGetPaintHwnd(out obj))
                {
                    Console.ForegroundColor = Colors.Error;
                    Console.WriteLine("[Program] ERROR: Can't find Paint Hwnd. Process aborted.");
                    return;
                }
                */
                Console.ForegroundColor = Colors.Message;
                Console.WriteLine("[Program] Will try again to get handle");
                return GetInstance(numRetries - 1);
            }
        }
    }
}
