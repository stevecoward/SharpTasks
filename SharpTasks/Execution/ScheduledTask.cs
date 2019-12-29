using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace SharpTasks.Execution
{
    public class ScheduledTask
    {
        private static string StringFormat = "{0,-100} {1,-10} {2,-20}";
        private static string ResultHeader = String.Format(StringFormat + Environment.NewLine, "Task", "Status", "Next Run");

        public sealed class ScheduledTaskResult
        {
            private string FormattedName;

            public string Name { get; set; } = "";
            public string Status { get; set; } = "";
            public string NextRun { get; set; } = "";

            public override string ToString()
            {
                FormattedName = Name.Length > 97 ? Name.Substring(0, 97) + "..." : Name;
                return String.Format(StringFormat, FormattedName, Status, NextRun);
            }
        }

        public static List<ScheduledTaskResult> GetScheduledTasks(string Folder)
        {
            List<ScheduledTaskResult> tasks = new List<ScheduledTaskResult>();

            using (Process process = new Process())
            {
                process.StartInfo.FileName = "schtasks.exe";
                process.StartInfo.Arguments = String.Format("/query /nh /fo csv /tn {0}", string.IsNullOrEmpty(Folder.Trim()) ? "\\" : Folder.Trim() + "\\");
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.Start();

                string processOutput = process.StandardOutput.ReadToEnd();
                string processStdErr = process.StandardError.ReadToEnd();

                foreach (string line in processOutput.Split('\n'))
                {
                    string[] lineData = line.Split(',');
                    if (lineData.Length > 2)
                    {
                        tasks.Add(new ScheduledTaskResult
                        {
                            Name = lineData[0].Replace("\"", ""),
                            NextRun = lineData[1].Replace("\"", ""),
                            Status = lineData[2].Replace("\"", "").Replace("\r", ""),
                        });
                    }
                }

                process.WaitForExit();
            }

            return tasks;
        }

        public static string GetResultsReport(List<ScheduledTaskResult> Tasks)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(ResultHeader);

            foreach (ScheduledTaskResult Task in Tasks)
            {
                builder.Append(Task.ToString() + Environment.NewLine);
            }

            return builder.ToString();
        }

        public static string GetResultsReport(ScheduledTaskResult Task)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(ResultHeader);
            builder.Append(Task.ToString());

            return builder.ToString();
        }

        public static string FilterTasksByName(List<ScheduledTaskResult> Tasks, string Name)
        {
            if (!string.IsNullOrEmpty(Name.Trim()))
            {
                ScheduledTaskResult foundTask = Tasks.Where(task => task.Name.Contains(Name)).FirstOrDefault();
                if (foundTask != null)
                    return GetResultsReport(foundTask);
            }

            return GetResultsReport(Tasks);
        }
    }
}