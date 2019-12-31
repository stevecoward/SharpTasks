using System.Diagnostics;

namespace SharpTasks.Generic
{
    /// <summary>
    ///     Container for any action related to creating/managing processes on a system.
    /// </summary>
    internal class Processes
    {
        /// <summary>
        ///     Create and execute a new process, capturing stdout/stderr.
        /// </summary>
        /// <param name="FileName">Absolute file path to program to execute.</param>
        /// <param name="Arguments">Additional arguments required for command to be run.</param>
        /// <param name="UseShellExecute">
        ///     Gets or sets a value indicating whether to use the operating system shell to start the
        ///     process.
        /// </param>
        /// <param name="RedirectStandardOutput">Direct process output to a variable that can be accessed/manipulated.</param>
        /// <param name="RedirectStandardError">
        ///     Direct process standard error output to a variable that can be
        ///     accessed/manipulated.
        /// </param>
        /// <returns>Results of the command run.</returns>
        public static string CreateProcess(string FileName, string Arguments, bool UseShellExecute = false,
            bool RedirectStandardOutput = true, bool RedirectStandardError = true)
        {
            using (var process = new Process())
            {
                process.StartInfo.FileName = FileName;
                process.StartInfo.Arguments = Arguments;
                process.StartInfo.UseShellExecute = UseShellExecute;
                process.StartInfo.RedirectStandardOutput = RedirectStandardOutput;
                process.StartInfo.RedirectStandardError = RedirectStandardError;
                process.Start();

                var processOutput = process.StandardOutput.ReadToEnd();
                var processStdErr = process.StandardError.ReadToEnd();

                process.WaitForExit();

                return processOutput;
            }
        }
    }
}