using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using SharpTasks.Generic;

namespace SharpTasks.Execution
{
    public class ScheduledTask
    {
        private static string StringFormat = "{0,-100} {1,-10} {2,-20}";
        private static string ResultHeader = String.Format(StringFormat + Environment.NewLine, "Task", "Status", "Next Run");
        private static List<ScheduleOptionProperties> ScheduleOptions = new List<ScheduleOptionProperties> {
            new ScheduleOptionProperties { Name = "MINUTE", MaximumValue = 1439 },
            new ScheduleOptionProperties { Name = "HOURLY", MaximumValue = 23 },
            new ScheduleOptionProperties { Name = "DAILY", MaximumValue = 365 },
            new ScheduleOptionProperties { Name = "WEEKLY", MaximumValue = 52 },
            new ScheduleOptionProperties { Name = "MONTHLY", MaximumValue = 12 },
            new ScheduleOptionProperties { Name = "ONCE", MaximumValue = 0 },
            new ScheduleOptionProperties { Name = "ONLOGON", MaximumValue = 0 },
            new ScheduleOptionProperties { Name = "ONIDLE", MaximumValue = 0 },
            new ScheduleOptionProperties { Name = "ONEVENT", MaximumValue = 0 },
        };

        public sealed class ScheduleOptionProperties
        {
            public string Name { get; set; } = "";
            public int MaximumValue { get; set; } = new int();
            public override string ToString()
            {
                return String.Format("Option: Name={0} | MaximumValue={1}", Name, MaximumValue);
            }
        }

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

            string output = Processes.CreateProcess("schtasks.exe", String.Format("/query /nh /fo csv /tn {0}", string.IsNullOrEmpty(Folder.Trim()) ? "\\" : Folder.Trim() + "\\"));
            foreach (string line in output.Split('\n'))
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
            ScheduledTaskResult foundTask = Tasks.Where(task => task.Name.Contains(Name)).FirstOrDefault();
            if (foundTask == null && !Name.Equals("*"))
                return String.Format("Unable to locate any scheduled tasks matching name: {0}", Name);

            if (foundTask != null)
                return GetResultsReport(foundTask);

            return GetResultsReport(Tasks);
        }

        public static string CreateScheduledTask(string Schedule, string Modifier, string Name, string Run)
        {
            ScheduleOptionProperties property = ScheduleOptions.Find(prop => prop.Name.ToLower().Equals(Schedule.ToLower()));
            if (property == null)
                return String.Format("Invalid value for 'Schedule': {0}", Schedule);

            if (int.Parse(Modifier) > property.MaximumValue)
                return "Modifier for task exceeds maximum value. " + property.ToString();

            string[] taskNameParts = Name.Split('\\');
            string[] taskFolderParts = new string[taskNameParts.Length - 1];
            string taskName = taskNameParts[taskNameParts.Length - 1];

            for (int i = 0; i < taskNameParts.Length - 1; i++)
                taskFolderParts[i] = taskNameParts[i];

            List<ScheduledTask.ScheduledTaskResult> tasks = ScheduledTask.GetScheduledTasks(String.Join("\\", taskFolderParts));
            if (tasks != null && ScheduledTask.FilterTasksByName(tasks, taskName).Contains("Unable to locate"))
            {
                string output = Processes.CreateProcess("schtasks.exe", String.Format("/create /sc {0} /mo {1} /tn \"{2}\" /tr \"{3}\"", property.Name, Modifier, String.Join("\\", taskFolderParts) + "\\" + taskName, Run));
                return (output.Contains("SUCCESS")) ? output : "Failed to create scheduled task";
            }

            return "Failed to create task. Task probably exists already.";
        }
    }
}