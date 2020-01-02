using System;
using System.IO;
using System.Linq;

namespace SharpTasks.Evasion
{
    public class Timestomp
    {
        private enum EditableTimestampFields { Modified, Accessed, Created, Changed, All }

        /// <summary>
        ///     Modifies a file's timestamp values
        /// </summary>
        /// <remarks>
        ///     User should be able to set:
        ///         - Modified Date
        ///         - Access Date
        ///         - Create/Change Timestamps
        /// </remarks>
        /// <param name="FileName">Path to targeted file.</param>
        /// <param name="WhichTimestamps">Field to select which timestamps to update</param>
        /// <param name="Timestamp">Timestamp value to use for file.</param>
        public static string SetFileTimestamp(string FileName, string WhichTimestamps, string Timestamp)
        {
            // Do some string manipulation to match user input to Title case for enum var
            var timestampSelection = WhichTimestamps.First().ToString().ToUpper() + WhichTimestamps.Substring(1);
            var timestampFieldOptions = Enum.GetNames(typeof(EditableTimestampFields))
                .Select(Field => Field.ToString())
                .ToArray();
            var timestamp = DateTime.Parse(Timestamp);
            var output = $"Updating '{timestampSelection}' property for file: {FileName}";

            if (!File.Exists(FileName))
                return $"File does not exist: {FileName}";

            if (!Enum.IsDefined(typeof(EditableTimestampFields), timestampSelection))
                return $"Timestamp field to edit is invalid. Choices are: [{string.Join(", ", timestampFieldOptions)}]";

            switch (Enum.Parse(typeof(EditableTimestampFields), timestampSelection))
            {
                case EditableTimestampFields.Accessed:
                    File.SetLastAccessTime(FileName, timestamp);
                    break;
                case EditableTimestampFields.Created:
                    File.SetCreationTime(FileName, timestamp);
                    break;
                case EditableTimestampFields.Modified:
                case EditableTimestampFields.Changed:
                    File.SetLastWriteTime(FileName, timestamp);
                    break;
                case EditableTimestampFields.All:
                    File.SetLastAccessTime(FileName, timestamp);
                    File.SetCreationTime(FileName, timestamp);
                    File.SetLastWriteTime(FileName, timestamp);
                    break;
            }

            return output;
        }
    }
}
