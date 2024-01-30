using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using CommandLine;
using Microsoft.Isam.Esent.Interop;
using Microsoft.Isam.Esent.Interop.Vista;

namespace dumpntds
{
    class Program
    {
        /// <summary>
        /// The columns from the "ntds" ESENT database, just the columns from the
        /// "datatable" table that are needed by the "ntdsxtract" script to parse
        /// the user details including password hashes
        /// </summary>
        private static List<string> userColumns = new List<string> { "DNT_col", "PDNT_col", "time_col",
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
        static void Main(string[] args)
        {
            var res = Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);
        }

        static void RunOptions(Options opts)
        {
            if (File.Exists(opts.Ntds) == false)
            {
                Console.WriteLine("ntds.dit file does not exist");
            }

            Api.JetSetSystemParameter(JET_INSTANCE.Nil, JET_SESID.Nil, JET_param.DatabasePageSize, (int)8192, null);

            using (var instance = new Instance("dumpntds"))
            {
                instance.Parameters.Recovery = false;
                instance.Init();

                using (var session = new Session(instance))
                {
                    JET_DBID dbid;

                    Api.JetAttachDatabase(session, opts.Ntds, AttachDatabaseGrbit.ReadOnly);
                    Api.JetOpenDatabase(session, opts.Ntds, null, out dbid, OpenDatabaseGrbit.ReadOnly);

                    ExportDataTable(session, dbid);
                    ExportLinkTable(session, dbid);
                }
            }
        }

        static void HandleParseError(IEnumerable<Error> errs)
        {
            // We can log it out but fo now, we will ignore.
        }

        private static void ExportDataTable(Session session, JET_DBID dbid)
        {
            // Extract and cache the columns from the "datatable" table. Note
            // that we are only interested in the columns needed for "ntdsextract"
            List<ColumnInfo> columns = new List<ColumnInfo>();
            foreach (ColumnInfo column in Api.GetTableColumns(session, dbid, "datatable"))
            {
                if (userColumns.Contains(column.Name) == false)
                {
                    continue;
                }
                columns.Add(column);
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter("datatable.csv"))
            using (var table = new Microsoft.Isam.Esent.Interop.Table(session, dbid, "datatable", OpenTableGrbit.ReadOnly))
            {
                // Write out the column headers
                int index = 0;
                foreach (string property in userColumns)
                {
                    index += 1;
                    file.Write(property);
                    if (index != userColumns.Count())
                    {
                        file.Write("\t");
                    }
                }
                file.WriteLine("");

                Api.JetSetTableSequential(session, table, SetTableSequentialGrbit.None);
                Api.MoveBeforeFirst(session, table);

                int currentRow = 0;
                string formattedData = "";
                while (Api.TryMoveNext(session, table))
                {
                    currentRow++;

                    dynamic data = new ExpandoObject();
                    IDictionary<string, object> obj = data;
                    foreach (ColumnInfo column in columns)
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
                    foreach (string property in userColumns)
                    {
                        index += 1;
                        if (obj.TryGetValue(property, out var val))
                        {
                            file.Write(val);
                        }
                        else
                        {
                            file.Write(string.Empty);
                        }
                        if (index != userColumns.Count())
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
            using (System.IO.StreamWriter file = new System.IO.StreamWriter("linktable.csv"))
            using (var table = new Microsoft.Isam.Esent.Interop.Table(session, dbid, "link_table", OpenTableGrbit.ReadOnly))
            {
                // Extract and cache the columns from the "link_table" table
                List<ColumnInfo> columns = new List<ColumnInfo>();
                foreach (ColumnInfo column in Api.GetTableColumns(session, dbid, "link_table"))
                {
                    columns.Add(column);
                }

                // Write out the column headers
                int index = 0;
                foreach (ColumnInfo column in columns)
                {
                    index += 1;
                    file.Write(column.Name);
                    if (index != columns.Count())
                    {
                        file.Write("\t");
                    }
                }
                file.WriteLine("");

                Api.JetSetTableSequential(session, table, SetTableSequentialGrbit.None);
                Api.MoveBeforeFirst(session, table);

                List<dynamic> temp = new List<dynamic>();

                int currentRow = 0;
                string formattedData = "";
                while (Api.TryMoveNext(session, table))
                {
                    currentRow++;

                    dynamic data = new ExpandoObject();
                    IDictionary<string, object> obj = data;
                    foreach (ColumnInfo column in columns)
                    {
                        formattedData = GetFormattedColumnData(session, table, column);
                        obj.Add(column.Name, formattedData.Replace("\0", string.Empty));
                    }

                    // Now write out each columns data
                    index = 0;
                    foreach (ColumnInfo column in columns)
                    {
                        index += 1;
                        file.Write(obj[column.Name]);
                        if (index != columns.Count())
                        {
                            file.Write("\t");
                        };
                    }
                    file.WriteLine("");
                }


                Api.JetResetTableSequential(session, table, ResetTableSequentialGrbit.None);
            }
        }

        /// <summary>
        /// Extracts the column data as the correct data type and formats appropriately
        /// </summary>
        /// <param name="session"></param>
        /// <param name="table"></param>
        /// <param name="columnInfo"></param>
        /// <returns></returns>
        public static string GetFormattedColumnData(Session session,
                                                    JET_TABLEID table,
                                                    ColumnInfo columnInfo)
        {
            try
            {
                string temp = "";

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
                        Encoding encoding = (columnInfo.Cp == JET_CP.Unicode) ? Encoding.Unicode : Encoding.ASCII;
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

                if (temp == null)
                {
                    return "";
                }

                return temp;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error formatting data: " + ex.ToString());
                return string.Empty;
            }
        }


        /// <summary>
        /// Return the string format of a byte array.
        /// </summary>
        /// <param name="data">The data to format.</param>
        /// <returns>A string representation of the data.</returns>
        public static string FormatBytes(byte[] data)
        {
            if (null == data)
            {
                return null;
            }

            var sb = new StringBuilder(data.Length * 2);
            foreach (byte b in data)
            {
                sb.AppendFormat("{0:x2}", b);
            }

            return sb.ToString();
        }
    }
}
