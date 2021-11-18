using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput;

namespace NeosZone.EmployeeOfTheMonth
{
    class Program
    {
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        private static IntPtr notepadHandle;
        private static InputSimulator _sim;

        static void Main(string[] args) => Refresh();

        private static void Refresh()
        {
            Console.WriteLine($"Welcome to Employee of the month \n Close this window to exit.");

            while (true)
            {
                _sim = new InputSimulator();

                //Set notepadc
                Process notepad = new Process();
                notepad.StartInfo.FileName = "notepad.exe";

                notepad.Start();
                notepad.WaitForInputIdle();
                notepadHandle = notepad.MainWindowHandle;

                Write("Hello,");
                MoveMouseRandom();
                _sim.Keyboard.TextEntry("\n");

                Write("Thanks for your email,");
                _sim.Keyboard.TextEntry("\n");
                Write("I will get to you asap");
                _sim.Keyboard.TextEntry("\n");
                _sim.Keyboard.TextEntry("\n");
                Write("Regards");
                MoveMouseRandom();

                notepad.Kill();

                var asd = new Random();
                var next = asd.Next(12000, 60000);
                Console.WriteLine($"Next in {next/1000} sec");
                Thread.Sleep(next);
            }
        }

        private static void Write(string text)
        {
            SetForegroundWindow(notepadHandle);


            foreach (var letter in text)
            {
                _sim.Keyboard.TextEntry(letter);
                Thread.Sleep(200);
            }

        }

        private static void MoveMouseRandom()
        {
            var counter = 0;

            _sim.Mouse.MoveMouseBy(-10000, -10000);
            _sim.Mouse.MoveMouseBy(200, 200);

            while (counter < 1000)
            {
                var xRandom = new Random();
                var yRandom = new Random();

                var y = yRandom.Next(0, 2);
                var x = xRandom.Next(0, 2);

                var delta = 1;
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);
                _sim.Mouse.MoveMouseBy(y == 1 ? delta : -delta, x == 1 ? delta : -delta);

                counter++;
            }
        }
    }
}
