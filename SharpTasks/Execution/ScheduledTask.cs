using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharpTasks.Generic;

namespace SharpTasks.Execution
{
    /// <summary>
    /// Holds methods for interacting with a Windows Scheduled Task.
    /// </summary>
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

        /// <summary>
        /// Object that holds the properties of a scheduled task time option.
        /// </summary>
        public sealed class ScheduleOptionProperties
        {
            public string Name { get; set; } = "";
            public int MaximumValue { get; set; } = new int();
            public override string ToString()
            {
                return String.Format("Option: Name={0} | MaximumValue={1}", Name, MaximumValue);
            }
        }

        /// <summary>
        /// Object that stores a scheduled task's name, status and next run time.
        /// </summary>
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

        /// <summary>
        /// Get all scheduled tasks inside a given <c>Folder</c>.
        /// </summary>
        /// <param name="Folder">Scheduled task folder name. Defaults to "\" if value is empty/null.</param>
        /// <returns>A list of all scheduled tasks found in <c>Folder</c>.</returns>
        public static List<ScheduledTaskResult> GetScheduledTasks(string Folder)
        {
            List<ScheduledTaskResult> tasks = new List<ScheduledTaskResult>();

            /*
             * I originally tried using the TaskScheduler lib on nuget, which works fine normally but cannot get it working through Covenant (possible lib mismatch/exceptions)
             * Unable to use the Interop.TaskScheduler library either due to issue with dependencies, mismatch between mscorlib2 and mscorlib4 required for library.
             * Also tried implementing a WMI query to fetch scheduled tasks, but it turns out the results of that particular WMI query only returns tasks created using `at`.
             * Using schtasks.exe is NOT opsec-safe as it will pop a console window during the query process. Not ideal, but dammit it had to be done.
             */
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

        /// <summary>
        /// Loop through a list of <c>ScheduledTaskResult</c> objects and provide the toString() output along with a table header.
        /// </summary>
        /// <param name="Tasks">A list of tasks in the form of a <c>ScheduledTaskResult</c>.</param>
        /// <returns>A simple string-formatted table of scheduled tasks.</returns>
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

        /// <summary>
        /// Provide a simple string-formtted table of a scheduled task.
        /// </summary>
        /// <param name="Task">A task in the form of a <c>ScheduledTaskResult</c>.</param>
        /// <returns>A simple string-formatted table of scheduled task.</returns>
        public static string GetResultsReport(ScheduledTaskResult Task)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(ResultHeader);
            builder.Append(Task.ToString());

            return builder.ToString();
        }

        /// <summary>
        /// For a list of scheduled tasks, find a task by a given <c>Name</c>.
        /// </summary>
        /// <param name="Tasks">A list of <c>ScheduledTaskResult</c> objects to filter through.</param>
        /// <param name="Name">The name of the scheduled task to search for.</param>
        /// <returns>Returns either a found scheduled task or the entire list of scheduled tasks.</returns>
        /// <todo>Change the behavior of outcome. Don't return a report of all tasks if the filtered result is blank.</todo>
        public static string FilterTasksByName(List<ScheduledTaskResult> Tasks, string Name)
        {
            ScheduledTaskResult foundTask = Tasks.Where(task => task.Name.Contains(Name)).FirstOrDefault();
            if (foundTask == null && !Name.Equals("*"))
                return String.Format("Unable to locate any scheduled tasks matching name: {0}", Name);

            if (foundTask != null)
                return GetResultsReport(foundTask);

            return GetResultsReport(Tasks);
        }

        /// <summary>
        /// Creates a new scheduled task
        /// </summary>
        /// <param name="Schedule">Timeframe in which a task is to be run.</param>
        /// <param name="Modifier">Integer value to modify the Scheduled.</param>
        /// <param name="Name">Name for the scheduled task.</param>
        /// <param name="Run">Executable path and parameters to run with scheduled task.</param>
        /// <returns>Outcome of task creation.</returns>
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