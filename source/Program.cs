﻿using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Text;
using CommandLine;
using Microsoft.Isam.Esent.Interop;
using Microsoft.Isam.Esent.Interop.Vista;

namespace dumpntds
{
    internal static class Program
    {
        /// <summary>
        /// The columns from the "ntds" ESENT database, just the columns from the
        /// "datatable" table that are needed by the "ntdsxtract" script to parse
        /// the user details including password hashes
        /// </summary>
        private static readonly List<string> userColumns = new List<string> { "DNT_col", "PDNT_col", "time_col",
            "Ancestors_col", "ATTb590606", "ATTm3", "ATTm589825", "ATTk589826", "ATTl131074", "ATTl131075",
            "ATTq131091", "ATTq131192", "OBJ_col", "ATTi131120", "ATTb590605", "ATTr589970", "ATTm590045",
            "ATTm590480", "ATTj590126", "ATTj589832", "ATTq589876", "ATTq591520", "ATTq589983", "ATTq589920",
            "ATTq589873", "ATTj589993", "ATTj589836", "ATTj589922", "ATTk589914", "ATTk589879", "ATTk589918",
            "ATTk589984", "ATTk591734", "ATTk36", "ATTk589949", "ATTj589993", "ATTm590443", "ATTm590187",
            "ATTm590188", "ATTm591788", "ATTk591823", "ATTk591822", "ATTk591789", "ATTi590943", "ATTk590689" };

        /// <summary>
        /// Application entry point
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args)
        {
            var res = Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);
        }

        internal static void HandleParseError(IEnumerable<Error> errs)
        {
            // We can log it out but fo now, we will ignore.
        }

        internal static void RunOptions(Options opts)
        {
            if (!File.Exists(opts.Ntds))
            {
                Console.WriteLine("ntds.dit file does not exist");
            }

            Api.JetSetSystemParameter(JET_INSTANCE.Nil, JET_SESID.Nil, JET_param.DatabasePageSize, 8192, null);

            using (var instance = new Instance("dumpntds"))
            {
                instance.Parameters.Recovery = false;
                instance.Init();

                using (var session = new Session(instance))
                {
                    Api.JetAttachDatabase(session, opts.Ntds, AttachDatabaseGrbit.ReadOnly);
                    Api.JetOpenDatabase(session, opts.Ntds, null, out var dbid, OpenDatabaseGrbit.ReadOnly);

                    ExportCsv(session, dbid);
                }
            }
        }

        private static void ExportCsv(Session session, JET_DBID dbid)
        {
            ExportDataTable(session, dbid);
            ExportLinkTable(session, dbid);
        }

        private static void ExportDataTable(Session session, JET_DBID dbid)
        {
            // Extract and cache the columns from the "datatable" table. Note
            // that we are only interested in the columns needed for "ntdsextract"
            var columns = new List<ColumnInfo>();
            foreach (var column in Api.GetTableColumns(session, dbid, "datatable"))
            {
                if (!userColumns.Contains(column.Name))
                {
                    continue;
                }
                columns.Add(column);
            }

            using (var file = new StreamWriter("datatable.csv"))
            using (var table = new Table(session, dbid, "datatable", OpenTableGrbit.ReadOnly))
            {
                // Write out the column headers
                var index = 0;
                foreach (var property in userColumns)
                {
                    index++;
                    file.Write(property);
                    if (index != userColumns.Count)
                    {
                        file.Write("\t");
                    }
                }
                file.WriteLine("");

                Api.JetSetTableSequential(session, table, SetTableSequentialGrbit.None);
                Api.MoveBeforeFirst(session, table);

                var currentRow = 0;
                var formattedData = "";
                while (Api.TryMoveNext(session, table))
                {
                    currentRow++;

                    IDictionary<string, object> obj = new ExpandoObject();
                    foreach (var column in columns)
                    {
                        formattedData = GetFormattedColumnData(session, table, column);
                        // The first row has a null "PDNT_col" value which causes issues with the "ntdsxtract" scripts.
                        // esedbexport seems to have some other value, esedbexport parses the data, rather than using the API
                        if (column.Name == "PDNT_col")
                        {
                            if (formattedData.Length == 0)
                            {
                                obj.Add(column.Name, "0");
                                continue;
                            }
                        }

                        obj.Add(column.Name, formattedData.Replace("\0", string.Empty));
                    }

                    // Now write out each columns data
                    index = 0;
                    foreach (var property in userColumns)
                    {
                        index++;
                        if (obj.TryGetValue(property, out var val))
                        {
                            file.Write(val);
                        }
                        else
                        {
                            file.Write(string.Empty);
                        }
                        if (index != userColumns.Count)
                        {
                            file.Write("\t");
                        }
                    }
                    file.WriteLine("");
                }

                Api.JetResetTableSequential(session, table, ResetTableSequentialGrbit.None);
            }
        }

        private static void ExportLinkTable(Session session, JET_DBID dbid)
        {
            using (var file = new StreamWriter("linktable.csv"))
            using (var table = new Table(session, dbid, "link_table", OpenTableGrbit.ReadOnly))
            {
                // Extract and cache the columns from the "link_table" table
                var columns = new List<ColumnInfo>(Api.GetTableColumns(session, dbid, "link_table"));

                // Write out the column headers
                var index = 0;
                foreach (var column in columns)
                {
                    index++;
                    file.Write(column.Name);
                    if (index != columns.Count)
                    {
                        file.Write("\t");
                    }
                }
                file.WriteLine("");

                Api.JetSetTableSequential(session, table, SetTableSequentialGrbit.None);
                Api.MoveBeforeFirst(session, table);

                var temp = new List<dynamic>();

                var currentRow = 0;
                var formattedData = "";
                while (Api.TryMoveNext(session, table))
                {
                    currentRow++;

                    IDictionary<string, object> obj = new ExpandoObject();
                    foreach (var column in columns)
                    {
                        formattedData = GetFormattedColumnData(session, table, column);
                        obj.Add(column.Name, formattedData.Replace("\0", string.Empty));
                    }

                    // Now write out each columns data
                    index = 0;
                    foreach (var column in columns)
                    {
                        index++;
                        file.Write(obj[column.Name]);
                        if (index != columns.Count)
                        {
                            file.Write("\t");
                        }
                    }
                    file.WriteLine(string.Empty);
                }

                Api.JetResetTableSequential(session, table, ResetTableSequentialGrbit.None);
            }
        }

        /// <summary>
        /// Return the string format of a byte array.
        /// </summary>
        /// <param name="data">The data to format.</param>
        /// <returns>A string representation of the data.</returns>
        private static string FormatBytes(byte[] data)
        {
            if (data == null)
            {
                return null;
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
        private static string GetFormattedColumnData(Session session,
                                                    JET_TABLEID table,
                                                    ColumnInfo columnInfo)
        {
            try
            {
                var temp = "";

                switch (columnInfo.Coltyp)
                {
                    case JET_coltyp.Bit:
                        temp = string.Format("{0}", Api.RetrieveColumnAsBoolean(session, table, columnInfo.Columnid));
                        break;
                    case VistaColtyp.LongLong:
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
                    case JET_coltyp.Binary:
                    case JET_coltyp.LongBinary:
                    default:
                        temp = FormatBytes(Api.RetrieveColumn(session, table, columnInfo.Columnid));
                        break;
                }

                return temp ?? string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error formatting data: {ex}");
                return string.Empty;
            }
        }
    }
}
