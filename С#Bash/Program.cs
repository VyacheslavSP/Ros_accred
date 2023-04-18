using System.Diagnostics;
using System;
using System.Threading;

var startInfo = new ProcessStartInfo()
{
    FileName = "cmd.exe",
    Arguments = @"/k ""C:\Users\VecheslavSP\Desktop\Python\Ros_accred\main.py"""//настраиваем консоль
              + @" && exit",//закрываем консоль
    UseShellExecute = true
};
Process.Start(startInfo);


System.Threading.Thread.Sleep(280000);