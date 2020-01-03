using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharpTasks.Generic;

namespace SharpTasks.Execution
{
    /// <summary>
    ///     Holds methods for interacting with a Windows Scheduled Task.
    /// </summary>
    public class ScheduledTask
    {
        private const string StringFormat = "{0,-100} {1,-10} {2,-20}";

        private static readonly string ResultHeader =
            string.Format(StringFormat + Environment.NewLine, "Task", "Status", "Next Run");

        private static readonly List<ScheduleOptionProperties> ScheduleOptions = new List<ScheduleOptionProperties>
        {
            new ScheduleOptionProperties {Name = "MINUTE", MaximumValue = 1439},
            new ScheduleOptionProperties {Name = "HOURLY", MaximumValue = 23},
            new ScheduleOptionProperties {Name = "DAILY", MaximumValue = 365},
            new ScheduleOptionProperties {Name = "WEEKLY", MaximumValue = 52},
            new ScheduleOptionProperties {Name = "MONTHLY", MaximumValue = 12},
            new ScheduleOptionProperties {Name = "ONCE", MaximumValue = 0},
            new ScheduleOptionProperties {Name = "ONLOGON", MaximumValue = 0},
            new ScheduleOptionProperties {Name = "ONIDLE", MaximumValue = 0},
            new ScheduleOptionProperties {Name = "ONEVENT", MaximumValue = 0}
        };

        /// <summary>
        ///     Runs a given scheduled task.
        /// </summary>
        /// <param name="Name">Full path and name of scheduled task.</param>
        /// <returns>String output with outcome of running scheduled task.</returns>
        public static string Run(string Name)
        {
            var explodedPath = ExplodePath(Name);
            if (!Exists(explodedPath.Folder, explodedPath.Name))
                return $"Unable to edit task: {explodedPath.Name}. Task does not exist at path: {explodedPath.Folder}";

            var output = Processes.CreateProcess("schtasks.exe", $"/run /tn \"{Name}\"");

            return output;
        }

        /// <summary>
        ///     Deletes a given scheduled task.
        /// </summary>
        /// <param name="Name">Full path and name of scheduled task.</param>
        /// <param name="Force">Toggles the forceful deletion of scheduled task.</param>
        /// <returns>String output with outcome of command output.</returns>
        public static string Delete(string Name)
        {
            var explodedPath = ExplodePath(Name);
            if (!Exists(explodedPath.Folder, explodedPath.Name))
                return $"Unable to edit task: {explodedPath.Name}. Task does not exist at path: {explodedPath.Folder}";

            var output = Processes.CreateProcess("schtasks.exe", $"/delete /tn \"{Name}\" /f");

            return output;
        }

        /// <summary>
        ///     Changes certain properties of a scheduled task.
        /// </summary>
        /// <param name="Name">Folder and name of scheduled task.</param>
        /// <param name="Toggle">Must be either 'Enable' or 'Disable'.</param>
        /// <param name="Run">Executable path and parameters to run with scheduled task.</param>
        /// <param name="RunAsUser">User to make changes to scheduled task as.</param>
        /// <param name="RunAsPassword">Password for RunAsUser value.</param>
        /// <returns>Various strings indicating outcome of change operation.</returns>
        public static string Edit(string Name, string Toggle, string Run, string RunAsUser, string RunAsPassword)
        {
            var explodedPath = ExplodePath(Name);
            if (!Exists(explodedPath.Folder, explodedPath.Name))
                return $"Unable to edit task: {explodedPath.Name}. Task does not exist at path: {explodedPath.Folder}";

            if (!(Toggle.ToLower() == "enable" || Toggle.ToLower() == "disable"))
                return "Invalid toggle option chosen. Acceptable values are 'Enable' or 'Disable'";

            var toggleChoice = (Toggle.ToLower() == "enable") ? "/enable" : "/disable";
            var output = Processes.CreateProcess("schtasks.exe",
                $"/change /tn \"{Name}\" /tr \"{Run}\" {toggleChoice} /ru {RunAsUser} /rp {RunAsPassword}");

            return output;
        }

        /// <summary>
        ///     Get all scheduled tasks inside a given <c>Folder</c>.
        /// </summary>
        /// <param name="Folder">Scheduled task folder name. Defaults to "\" if value is empty/null.</param>
        /// <returns>A list of all scheduled tasks found in <c>Folder</c>.</returns>
        public static List<ScheduledTaskResult> Get(string Folder)
        {
            var tasks = new List<ScheduledTaskResult>();

            /*
             * I originally tried using the TaskScheduler lib on nuget, which works fine normally but cannot get it working through Covenant (possible lib mismatch/exceptions)
             * Unable to use the Interop.TaskScheduler library either due to issue with dependencies, mismatch between mscorlib2 and mscorlib4 required for library.
             * Also tried implementing a WMI query to fetch scheduled tasks, but it turns out the results of that particular WMI query only returns tasks created using `at`.
             * Using schtasks.exe is NOT opsec-safe as it will pop a console window during the query process. Not ideal, but dammit it had to be done.
             */
            var output = Processes.CreateProcess("schtasks.exe",
                $"/query /nh /fo csv /tn {(string.IsNullOrEmpty(Folder.Trim()) ? "\\" : Folder.Trim() + "\\")}");
            foreach (var line in output.Split('\n'))
            {
                var lineData = line.Split(',');
                if (lineData.Length > 2)
                    tasks.Add(new ScheduledTaskResult
                    {
                        Name = lineData[0].Replace("\"", ""),
                        NextRun = lineData[1].Replace("\"", ""),
                        Status = lineData[2].Replace("\"", "").Replace("\r", "")
                    });
            }

            return tasks;
        }

        /// <summary>
        ///     Creates a new scheduled task
        /// </summary>
        /// <param name="Schedule">Time frame in which a task is to be run.</param>
        /// <param name="Modifier">Integer value to modify the Scheduled.</param>
        /// <param name="Name">Name for the scheduled task.</param>
        /// <param name="Run">Executable path and parameters to run with scheduled task.</param>
        /// <returns>Outcome of task creation.</returns>
        public static string Create(string Schedule, string Modifier, string Name, string Run)
        {
            var property = ScheduleOptions.Find(Prop => Prop.Name.ToLower().Equals(Schedule.ToLower()));
            if (property == null)
                return $"Invalid value for 'Schedule': {Schedule}";

            if (int.Parse(Modifier) > property.MaximumValue)
                return "Modifier for task exceeds maximum value. " + property;

            var explodedPath = ExplodePath(Name);

            var tasks = Get(explodedPath.Folder);
            if (tasks == null || !FilterByName(tasks, explodedPath.Name).Contains("Unable to locate"))
                return "Failed to create task. Task probably exists already.";

            if (Exists(explodedPath.Folder, explodedPath.Name))
                return "Failed to create task. Task probably exists already.";

            var output = Processes.CreateProcess("schtasks.exe",
                $"/create /sc {property.Name} /mo {Modifier} /tn \"{explodedPath.Folder + "\\" + explodedPath.Name}\" /tr \"{Run}\"");
            return output.Contains("SUCCESS") ? output : "Failed to create scheduled task";
        }

        /// <summary>
        ///     Checks if a scheduled task folder exists.
        /// </summary>
        /// <param name="Path">Full path of scheduled task folder</param>
        /// <returns><c>true</c> or <c>false</c> if path exists.</returns>
        private static bool Exists(string Path)
        {
            return Get(Path).Count > 0;
        }

        /// <summary>
        ///     Checks if a scheduled task exists in a folder.
        /// </summary>
        /// <param name="Path">Full path of scheduled task folder</param>
        /// <param name="Name">Name of scheduled task</param>
        /// <returns><c>true</c> or <c>false</c> if path and task exist.</returns>
        private static bool Exists(string Path, string Name)
        {
            var tasks = Get(Path);
            return (tasks.Count > 0 && !FilterByName(tasks, Name).Contains("Unable to locate"));
        }

        /// <summary>
        ///     Helper method to separate a task name from its folder path.
        /// </summary>
        /// <param name="FullPath">Absolute path to scheduled task including task name.</param>
        /// <returns>Object consisting of scheduled task folder and name.</returns>
        private static ExplodedPath ExplodePath(string FullPath)
        {
            var pathParts = FullPath.Split('\\');
            var taskFolderParts = new string[pathParts.Length - 1];
            var taskName = pathParts[pathParts.Length - 1];

            for (var i = 0; i < pathParts.Length - 1; i++)
                taskFolderParts[i] = pathParts[i];

            return new ExplodedPath { Folder = string.Join("\\", taskFolderParts), Name = taskName };
        }

        /// <summary>
        ///     Loop through a list of <c>ScheduledTaskResult</c> objects and provide the toString() output along with a table
        ///     header.
        /// </summary>
        /// <param name="Tasks">A list of tasks in the form of a <c>ScheduledTaskResult</c>.</param>
        /// <returns>A simple string-formatted table of scheduled tasks.</returns>
        public static string GetResultsReport(IEnumerable<ScheduledTaskResult> Tasks)
        {
            var builder = new StringBuilder();
            builder.Append(ResultHeader);

            foreach (var task in Tasks) builder.Append(task + Environment.NewLine);

            return builder.ToString();
        }

        /// <summary>
        ///     Provide a simple string-formatted table of a scheduled task.
        /// </summary>
        /// <param name="Task">A task in the form of a <c>ScheduledTaskResult</c>.</param>
        /// <returns>A simple string-formatted table of scheduled task.</returns>
        public static string GetResultsReport(ScheduledTaskResult Task)
        {
            var builder = new StringBuilder();
            builder.Append(ResultHeader);
            builder.Append(Task);

            return builder.ToString();
        }

        /// <summary>
        ///     For a list of scheduled tasks, find a task by a given <c>Name</c>.
        /// </summary>
        /// <param name="Tasks">A list of <c>ScheduledTaskResult</c> objects to filter through.</param>
        /// <param name="Name">The name of the scheduled task to search for.</param>
        /// <returns>Returns either a found scheduled task or the entire list of scheduled tasks.</returns>
        /// <todo>Change the behavior of outcome. Don't return a report of all tasks if the filtered result is blank.</todo>
        public static string FilterByName(List<ScheduledTaskResult> Tasks, string Name)
        {
            var foundTask = Tasks.FirstOrDefault(Task => Task.Name.Contains(Name));
            if (foundTask == null && !Name.Equals("*"))
                return $"Unable to locate any scheduled tasks matching name: {Name}";

            return foundTask != null ? GetResultsReport(foundTask) : GetResultsReport(Tasks);
        }

        /// <summary>
        ///     Object that holds the properties of a scheduled task time option.
        /// </summary>
        public sealed class ScheduleOptionProperties
        {
            /// <summary>
            ///     Name of the Scheduled Task option.
            /// </summary>
            public string Name { get; set; } = "";

            /// <summary>
            ///     Maximum value of the modifier for the Scheduled Task option.
            /// </summary>
            public int MaximumValue { get; set; }

            /// <summary>
            ///     Override method to return string representation of a <c>ScheduleOptionProperties</c> object.
            /// </summary>
            /// <returns>String representation of <c>ScheduleOptionProperties</c> object</returns>
            public override string ToString()
            {
                return $"Option: Name={Name} | MaximumValue={MaximumValue}";
            }
        }

        /// <summary>
        ///     Object that stores a scheduled task's name, status and next run time.
        /// </summary>
        public sealed class ScheduledTaskResult
        {
            private string _formattedName;

            /// <summary>
            ///     Name of the scheduled task.
            /// </summary>
            public string Name { get; set; } = "";

            /// <summary>
            ///     Current status of the scheduled task.
            /// </summary>
            public string Status { get; set; } = "";

            /// <summary>
            ///     Date and time the scheduled task is due to run next.
            /// </summary>
            public string NextRun { get; set; } = "";

            /// <summary>
            ///     Override method to return string representation of a <c>ScheduledTaskResult</c> object.
            /// </summary>
            /// <returns>String representation of a <c>ScheduledTaskResult</c> object.</returns>
            public override string ToString()
            {
                _formattedName = Name.Length > 97 ? Name.Substring(0, 97) + "..." : Name;
                return string.Format(StringFormat, _formattedName, Status, NextRun);
            }
        }

        private sealed class ExplodedPath
        {
            public string Folder { get; set; } = "";
            
            public string Name { get; set; } = "";
        }
    }
}