using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using PlannerExAndImport.JSON;

namespace PlannerExAndImport
{
    public partial class Planner
    {
        /// <summary>
        /// Imports task from a spreadsheet.
        /// </summary>
        public static void ImportBulk()
        {
            var fileName = Program.GetInput("Please enter the name of the input file (*.xlsx): ");
            if (string.IsNullOrEmpty(fileName))
            {
                Console.WriteLine("No input file specified");
                return;
            }

            var allowMultiSelect = false;
            Plan[] plans = SelectPlan(allowMultiSelect);
            if (!allowMultiSelect && plans.Length != 1)
            {
                Console.WriteLine("You must select a plan");
                return;
            }

            var plan = plans.First();
            var bucket = SelectBucket(plan);
            if (bucket == null)
            {
                Console.WriteLine("You must select a bucket");
                return;
            }

            var tasks = ReadTasks(fileName)
                .OrderBy(x => x.OrderHint)
                .ToList();

            using (var httpClient = PreparePlannerClient())
            {
                SaveTasks(plan.Id, bucket.Id, false, tasks, httpClient);
            }

            Console.WriteLine("Import is done");
        }

        /// <summary>
        /// Reads the task from a spreadsheet.
        /// </summary>
        /// <returns>An enumerable of tasks.</returns>
        private static IEnumerable<PlannerTask> ReadTasks(string fileName)
        {
            using var reader = new SpreadsheetReader(new FileStream(fileName, FileMode.Open));

            while (true)
            {
                var task = reader.Read();
                if (task == null)
                {
                    // Done
                    break;
                }

                yield return task;
            }
        }
    }

    /// <summary>
    /// A task reader that reads from a spreadsheet (XLSX) file.
    /// </summary>
    public class SpreadsheetReader : IDisposable
    {
        /// <summary>
        /// The underlying stream.
        /// </summary>
        private readonly Stream _stream;

        /// <summary>
        /// The workbook.
        /// </summary>
        private readonly IXLWorkbook _workbook;

        /// <summary>
        /// The sheet.
        /// </summary>
        private readonly IXLWorksheet _worksheet;

        /// <summary>
        /// The current row.
        /// </summary>
        private IXLRow _currentRow;

        /// <summary>
        /// Creates a new <see cref="SpreadsheetReader"/>.
        /// </summary>
        /// <param name="stream">The underlying stream to read from.</param>
        public SpreadsheetReader(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));

            // Open document
            _workbook = new XLWorkbook(stream);
            _worksheet = _workbook.Worksheets.First();
            _currentRow = _worksheet.FirstRowUsed();
        }

        /// <summary>
        /// Reads a value from the spreadsheet.
        /// </summary>
        /// <returns>The value that was read.</returns>
        public PlannerTask Read()
        {
            while (!_currentRow.IsEmpty())
            {
                /*
                 * Percentage complete (0-100)
                 *   todo:          0
                 *   in progress:  50
                 *   completed:   100
                 * Task priority (0-10)
                 *   urgent:    0, 1*
                 *   important: 2, 3*, 4
                 *   medium:    5*, 6, 7
                 *   low:       8, 9*, 10
                */

                PlannerTask task;

                try
                {
                    var title = _currentRow.Cell(1).GetValue<string>();
                    var percentComplete = _currentRow.Cell(2).GetValue<int>();
                    var priority = _currentRow.Cell(3).GetValue<int>();
                    var orderHint = _currentRow.Cell(4).GetValue<string>();
                    var description = _currentRow.Cell(5).GetValue<string>();

                    task = new PlannerTask
                    {
                        Title = title,
                        PercentComplete = percentComplete,
                        Priority = priority,
                        OrderHint = orderHint,
                        TaskDetail = new TaskDetailResponse
                        {
                            Description = description,
                            Checklist = new Dictionary<string, Checklist>(),
                            References = new Dictionary<string, Reference>(),
                        },
                    };
                }
                catch
                {
                    // Skip row
                    continue;
                }
                finally
                {
                    GoToNextRow();
                }

                return task;
            }

            // End of file
            return null;
        }

        /// <summary>
        /// Advances to the next row.
        /// </summary>
        private void GoToNextRow()
        {
            _currentRow = _currentRow.RowBelow();
        }

        /// <summary>
        /// Releases all resources used by the <see cref="SpreadsheetReader"/> object.
        /// </summary>
        public void Dispose()
        {
            _workbook.Dispose();
            _stream.Dispose();
        }
    }
}