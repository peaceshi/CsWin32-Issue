using System;
using System.IO;

using Windows.Win32;
using Windows.Win32.System.Com;
using Windows.Win32.UI.Shell;

namespace CsWin32_Issue;

public static class Program
{
    private static void Main()
    {
        Console.WriteLine("Hello, World!");

        Directory.CreateDirectory("./out");

        var path = Path.GetFullPath("./out");
        //create a zip by yourself
        var path1 = Path.GetFullPath("./test.zip");

        PInvoke.CoCreateInstance(
            typeof(Shell).GUID,
            null,
            CLSCTX.CLSCTX_INPROC_SERVER,
            out IShellDispatch shell).ThrowOnFailure();
        // Extract zip to dir.
        shell.NameSpace(path).CopyHere(shell.NameSpace(path1).Items(), 16);
    }
}