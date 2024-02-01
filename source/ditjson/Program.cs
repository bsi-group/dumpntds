using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;
using System.Text.Json;
using CommandLine;
using Microsoft.Isam.Esent.Interop;
using Microsoft.Isam.Esent.Interop.Vista;
using Microsoft.Isam.Esent.Interop.Windows10;

namespace ditjson
{
    internal static class Program
    {
        private static readonly JsonSerializerOptions options = new()
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.Strict,
            WriteIndented = true,
        };

        /// <summary>
        ///     Application entry point
        /// </summary>
        /// <param name="args"></param>
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(Options))]
        [RequiresUnreferencedCode("Calls ditjson.Program.RunOptions(Options)")]
        public static void Main(string[] args) => _ = Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);

        /// <summary>
        ///     Runs the incorrect parameter actions
        /// </summary>
        /// <param name="errs"></param>
        internal static void HandleParseError(IEnumerable<Error> errs) => Console.WriteLine("Check the parameters and retry.");

        /// <summary>
        ///     Runs the happy path code here.
        /// </summary>
        /// <param name="opts">Parameters as Options</param>
        /// <exception cref="NtdsException"></exception>
        /// <exception cref="FormatException"></exception>
        /// <exception cref="OverflowException"></exception>
        [RequiresUnreferencedCode("Calls ditjson.Program.ExportJson(Session, JET_DBID)")]
        internal static void RunOptions(Options opts)
        {
            if (!File.Exists(opts.Ntds))
            {
                Console.WriteLine($"ntds.dit file does not exist in the path {opts.Ntds}");
            }

            Api.JetSetSystemParameter(JET_INSTANCE.Nil, JET_SESID.Nil, JET_param.DatabasePageSize, 8192, null);

            using var instance = new Instance("ditjson");
            instance.Parameters.Recovery = false;
            instance.Init();

            using var session = new Session(instance);
            Api.JetAttachDatabase(session, opts.Ntds, AttachDatabaseGrbit.ReadOnly);
            Api.JetOpenDatabase(session, opts.Ntds, null, out var dbid, OpenDatabaseGrbit.ReadOnly);

            var ntdsDictionary = new Dictionary<string, object>
            {
                ["datatable"] = TableToList(session, dbid, "datatable"),
                ["linktable"] = TableToList(session, dbid, "link_table")
            };

            string json;
            try
            {
                json = JsonSerializer.Serialize(ntdsDictionary, options);
            }
            catch (NotSupportedException ex)
            {
                throw new NtdsException("Failed to serialize to JSON.", ex);
            }

            try
            {
                File.WriteAllText("ntds.json", json);
            }
            catch (Exception ex)
            {
                throw new NtdsException("Failed to write to JSON to file.", ex);
            }
        }

        /// <summary>
        /// Return the string format of a byte array.
        /// </summary>
        /// <param name="data">The data to format.</param>
        /// <returns>A string representation of the data.</returns>
        /// <exception cref="FormatException"></exception>
        /// <exception cref="OverflowException"></exception>
        private static string FormatBytes(byte[] data)
        {
            if (data == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder(data.Length * 2);
            foreach (var b in data)
            {
                sb.AppendFormat("{0:x2}", b);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Extracts the column data as the correct data type and formats appropriately
        /// </summary>
        /// <param name="session"></param>
        /// <param name="table"></param>
        /// <param name="columnInfo"></param>
        /// <returns></returns>
        /// <exception cref="NtdsException"></exception>
        /// <exception cref="FormatException"></exception>
        /// <exception cref="OverflowException"></exception>
        private static string GetFormattedValue(Session session,
                                                    JET_TABLEID table,
                                                    ColumnInfo columnInfo)
        {
            var temp = string.Empty;

            // Edge case: link_data_v2 column cannot be retreived properly
            // None of the Api.RetrieveColumnXXX commands can process this column.
            if (columnInfo.Name.Equals("link_data_v2"))
            {
                return temp;
            }

            switch (columnInfo.Coltyp)
            {
                case JET_coltyp.Bit:
                    temp = string.Format("{0}", Api.RetrieveColumnAsBoolean(session, table, columnInfo.Columnid));
                    break;
                case VistaColtyp.LongLong:
                case Windows10Coltyp.UnsignedLongLong:
                case JET_coltyp.Currency:
                    temp = string.Format("{0}", Api.RetrieveColumnAsInt64(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.IEEEDouble:
                    temp = string.Format("{0}", Api.RetrieveColumnAsDouble(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.IEEESingle:
                    temp = string.Format("{0}", Api.RetrieveColumnAsFloat(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.Long:
                    temp = string.Format("{0}", Api.RetrieveColumnAsInt32(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.Text:
                case JET_coltyp.LongText:
                    var encoding = (columnInfo.Cp == JET_CP.Unicode) ? Encoding.Unicode : Encoding.ASCII;
                    temp = string.Format("{0}", Api.RetrieveColumnAsString(session, table, columnInfo.Columnid, encoding));
                    break;
                case JET_coltyp.Short:
                    temp = string.Format("{0}", Api.RetrieveColumnAsInt16(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.UnsignedByte:
                    temp = string.Format("{0}", Api.RetrieveColumnAsByte(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.DateTime:
                    temp = string.Format("{0}", Api.RetrieveColumnAsDateTime(session, table, columnInfo.Columnid));
                    break;
                case VistaColtyp.UnsignedShort:
                    temp = string.Format("{0}", Api.RetrieveColumnAsUInt16(session, table, columnInfo.Columnid));
                    break;
                case VistaColtyp.UnsignedLong:
                    temp = string.Format("{0}", Api.RetrieveColumnAsUInt32(session, table, columnInfo.Columnid));
                    break;
                case VistaColtyp.GUID:
                    temp = string.Format("{0}", Api.RetrieveColumnAsGuid(session, table, columnInfo.Columnid));
                    break;
                case JET_coltyp.Nil:
                    break;
                case JET_coltyp.Binary:
                case JET_coltyp.LongBinary:
                    var columnBytes = Api.RetrieveColumn(session, table, columnInfo.Columnid);
                    if (columnBytes != null)
                    {
                        temp = FormatBytes(columnBytes);
                    }
                    break;

                default:
                    throw new NtdsException($"Unhandled column type {columnInfo.Coltyp} for {columnInfo.Columnid}");
            }

            return temp.Replace("\0", string.Empty);
        }

        /// <summary>
        ///     Export table as a <see cref="List{Dictionary{string, object}}"/>
        /// </summary>
        /// <param name="session">ESENT Session</param>
        /// <param name="dbid">Handle to the database</param>
        /// <returns>A <see cref="List{Dictionary{string, object}}"/> containing table data</returns>
        /// <exception cref="NtdsException"></exception>
        /// <exception cref="FormatException"></exception>
        /// <exception cref="OverflowException"></exception>
        private static List<IDictionary<string, object>> TableToList(Session session, JET_DBID dbid, string tableName)
        {
            var linktableValues = new List<IDictionary<string, object>>();
            var columns = new List<ColumnInfo>(Api.GetTableColumns(session, dbid, tableName));

            using (var table = new Table(session, dbid, tableName, OpenTableGrbit.ReadOnly))
            {
                Api.JetSetTableSequential(session, table, SetTableSequentialGrbit.None);
                Api.MoveBeforeFirst(session, table);

                var formattedData = string.Empty;
                while (Api.TryMoveNext(session, table))
                {
                    var obj = new Dictionary<string, object>();
                    foreach (var column in columns)
                    {
                        formattedData = GetFormattedValue(session, table, column);
                        var cellValue = formattedData;
                        // Ignore emptry or null values
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            obj.Add(column.Name, cellValue);
                        }
                    }

                    linktableValues.Add(obj);
                }

                Api.JetResetTableSequential(session, table, ResetTableSequentialGrbit.None);
            }

            return linktableValues;
        }
    }
}
