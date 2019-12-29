using System.Diagnostics;

namespace SharpTasks.Generic
{
    class Processes
    {
        public static string CreateProcess(string FileName, string Arguments, bool UseShellExecute = false, bool RedirectStandardOutput = true, bool RedirectStandardError = true)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = FileName;
                process.StartInfo.Arguments = Arguments;
                process.StartInfo.UseShellExecute = UseShellExecute;
                process.StartInfo.RedirectStandardOutput = RedirectStandardOutput;
                process.StartInfo.RedirectStandardError = RedirectStandardError;
                process.Start();

                string processOutput = process.StandardOutput.ReadToEnd();
                string processStdErr = process.StandardError.ReadToEnd();

                process.WaitForExit();

                return processOutput;
            }
        }
    }
}
