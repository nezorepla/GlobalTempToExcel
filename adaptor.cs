using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;


namespace adaptor
{
    class Program
    {
        public static SqlConnection baglanti;

        public static void exec(string q)
        {
            SqlCommand cmd = new SqlCommand(q, baglanti);

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }

            cmd.ExecuteNonQuery();
            //      baglanti.Close();
        }

        public static DataTable Getdata(string query)
        {
            // SqlConnection conn = new SqlConnection(GetConnStr(pConnKey));
            SqlDataAdapter da;
            DataTable dt = new DataTable();
            try
            {
                da = new SqlDataAdapter(query, baglanti);
                if (baglanti.State == ConnectionState.Closed)
                    baglanti.Open();
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                //   conn.Close();
            }
            return dt;
        }
        static void Main(string[] args)
        {
            baglanti = new SqlConnection("Data Source=XXX; Initial Catalog=dw_production; Integrated Security=true");

          //  exec("select 'a' a, '2' b, 3 c into ##deneme");
            DataTable dt = Getdata("select  * from ##DEVIR");

            create(dt);

            //writer.WriteCell(1, 0, "int");
            //writer.WriteCell(1, 1, 10);
            //writer.WriteCell(2, 0, "double");
            //writer.WriteCell(2, 1, 1.5);
            //writer.WriteCell(3, 0, "empty");
            //writer.WriteCell(3, 1);

        }
        public static void create(DataTable dt)
        {
            FileStream stream = new FileStream("DEVIR.xls", FileMode.OpenOrCreate);
            ExcelWriter writer = new ExcelWriter(stream);
            writer.BeginWrite();

          for (int i = 0; i < dt.Columns.Count; i++)
            {
                string name = dt.Columns[i].ColumnName.ToString();
                writer.WriteCell(0, i, name);

            }
			
		 /*  */
				
				
				
				 for (int r = 0; r < dt.Rows.Count; r++)
                    {
 //string TempStr = HeadStr;
 //Sb.Append("{");
  for (int c = 0; c < dt.Columns.Count; c++)
                        {
							      writer.WriteCell(r+1, c, dt.Rows[r][c].ToString());
 //  TempStr = TempStr.Replace("<br>", Environment.NewLine).Replace(Dt.Columns[j] + j.ToString() + "¾", Dt.Rows[r][c].ToString());
 }
 //Sb.Append(TempStr + "},");
 }
				
				
				
				
				
            writer.EndWrite();
            stream.Close();
        }

        /// <summary>
        /// Produces Excel file without using Excel
        /// </summary>
        public class ExcelWriter
        {
            private Stream stream;
            private BinaryWriter writer;

            private ushort[] clBegin = { 0x0809, 8, 0, 0x10, 0, 0 };
            private ushort[] clEnd = { 0x0A, 00 };


            private void WriteUshortArray(ushort[] value)
            {
                for (int i = 0; i < value.Length; i++)
                    writer.Write(value[i]);
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
            /// </summary>
            /// <param name="stream">The stream.</param>
            public ExcelWriter(Stream stream)
            {
                this.stream = stream;
                writer = new BinaryWriter(stream);
            }

            /// <summary>
            /// Writes the text cell value.
            /// </summary>
            /// <param name="row">The row.</param>
            /// <param name="col">The col.</param>
            /// <param name="value">The string value.</param>
            public void WriteCell(int row, int col, string value)
            {
                ushort[] clData = { 0x0204, 0, 0, 0, 0, 0 };
                int iLen = value.Length;
                byte[] plainText = Encoding.ASCII.GetBytes(value);
                clData[1] = (ushort)(8 + iLen);
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                clData[5] = (ushort)iLen;
                WriteUshortArray(clData);
                writer.Write(plainText);
            }

            /// <summary>
            /// Writes the integer cell value.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            /// <param name="value">The value.</param>
            public void WriteCell(int row, int col, int value)
            {
                ushort[] clData = { 0x027E, 10, 0, 0, 0 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
                int iValue = (value << 2) | 2;
                writer.Write(iValue);
            }

            /// <summary>
            /// Writes the double cell value.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            /// <param name="value">The value.</param>
            public void WriteCell(int row, int col, double value)
            {
                ushort[] clData = { 0x0203, 14, 0, 0, 0 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
                writer.Write(value);
            }

            /// <summary>
            /// Writes the empty cell.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            public void WriteCell(int row, int col)
            {
                ushort[] clData = { 0x0201, 6, 0, 0, 0x17 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
            }

            /// <summary>
            /// Must be called once for creating XLS file header
            /// </summary>
            public void BeginWrite()
            {
                WriteUshortArray(clBegin);
            }

            /// <summary>
            /// Ends the writing operation, but do not close the stream
            /// </summary>
            public void EndWrite()
            {
                WriteUshortArray(clEnd);
                writer.Flush();
            }
        }

       
    }
}
