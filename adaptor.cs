using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Text.RegularExpressions; 
using System.Diagnostics;

namespace adaptor
{
    class Program
    {
        public static SqlConnection baglanti;

        public static void exec(string q,int hata_bas)
        {
            SqlCommand cmd = new SqlCommand(q, baglanti);
cmd.CommandTimeout = 300;  

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
   try
            {
            cmd.ExecuteNonQuery();
            //      baglanti.Close();
    }
            catch (Exception ex)
            {
				if(hata_bas>0){
           Console.WriteLine("Hata");
		   Console.WriteLine(ex.ToString());
           Console.WriteLine("-------");
           Console.WriteLine(q);
		//   Console.ReadLine();
				}
            }
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

public static string dtGlobal;
 public static string Fx_into(	string query) {
	
	string q= query.ToUpper();
	q=q.Replace("  ", " ");
	q=q.Replace("  ", " ");
	q=q.Trim();
	q=q.Replace("FROM", " FROM");
	q=q.Replace("DROP TABLE ", "DROP_TABLE_");
	//q=q.Replace("USE ", "USE_");
	q=q.Replace(System.Environment.NewLine,"");
 q=q.Replace("DROP_TABLE_","INTO DROP_TABLE_");
	//q=q.Replace("USE_","INTO U");
	string rv= Left(vericek(q, "INTO ", " ").Trim(),40);

if(q.IndexOf("DROP ")>0)
{rv=q.Replace("INTO ", "");}
	
	
	return rv;
}
    public static string vericek(string StrData, string StrBas, string StrSon)
        {

            try
            {

                int IntBas = StrData.IndexOf(StrBas) + StrBas.Length;

                int IntSon = StrData.IndexOf(StrSon, IntBas + 1);

                return StrData.Substring(IntBas, IntSon - IntBas);

            }

            catch
            {

                return "";

            }

        }

 static void Main(string[] args)
        {
            baglanti = new SqlConnection("Data Source=10.180.20.31; Initial Catalog=dw_production; Integrated Security=true");

//	  string filepath = "d://ConvertedFile.csv";
//      DataTable res = ConvertCSVtoDataTable(filepath);
	  
	  
	   var dateAndTime_Glb = DateTime.Now;   
	   dtGlobal =dateAndTime_Glb.ToString();
	   
 Console.WriteLine(dtGlobal);
 
 

islem("PK_HESAP_DEVIR");
islem("PK_YAPILANDIRMALAR");
islem("PK_PERFORMANS");
islem("PK_PCSM");
islem("PK_DASHBOARD");

Process.Start(@"C:\Users\U05180\Desktop\kapakingen\DATAMART_MAIL.vbs");
	
System.Threading.Thread.Sleep(-1);
Console.ReadLine();
	}
		

		
    public static string Left(string gelen, int maxLength)
    {
        if (string.IsNullOrEmpty(gelen)) return gelen;
        maxLength = Math.Abs(maxLength);

        return ( gelen.Length <= maxLength 
               ? gelen 
               : gelen.Substring(0, maxLength)
               );
    }
	
public static void islem(string isim){

Console.WriteLine(isim+" Basladi-->"+DateTime.Now.ToString());


try{


string s_p=@"C:\Users\U05180\Desktop\sql\jobs\"+isim+".sql";

string text;
var fileStream = new FileStream(s_p, FileMode.Open, FileAccess.Read);
using (var streamReader = new StreamReader(fileStream, Encoding.Default))
{
	text = streamReader.ReadToEnd();
}


string[] ayir = text.Split(';');

int  payda=	ayir.Length;
int pay=1;

foreach (string parca in ayir)
{	
int hata_bas= 1;
string fxi=	Fx_into(parca);
if(Left(fxi.Replace(" ", ""),4) == "DROP"||fxi.IndexOf("DROP")>0) { 
hata_bas=0;
}	
			
Console.WriteLine(pay.ToString() +"/"+ payda.ToString() +"---> "+fxi);

if(parca.Length>5){

			exec(parca,hata_bas);
			}
		 

		 	
			//	if(fxi.Substring(0, 4) =="DROP") { hata_bas=0;}

		pay++;
 	 }			  



	




DataTable dt = Getdata("SELECT * FROM ##"+isim);

Console.WriteLine("##"+isim+" Tablosu tamamlandi"+DateTime.Now.ToString());

// DataTableToCSV(dt,isim+".csv");
mail();

}
catch(Exception e) {

Console.WriteLine("main class");
Console.WriteLine(e.ToString());
//Console.ReadLine();
}
}
		/*
		
		
		
public static void CSVCreate(DataTable dt){
	
	
	StringBuilder sb = new StringBuilder(); 

IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                  Select(column => column.ColumnName);
sb.AppendLine(string.Join("|", columnNames));

foreach (DataRow row in dt.Rows)
{
    IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
    sb.AppendLine(string.Join("|", fields));
}

File.WriteAllText("YAPILANDIRMALAR.csv", sb.ToString());
}

*/


public static void DataTableToCSV(DataTable dataTable, string filePath){


string pth=@"\\hegel\Perakende Krediler2\Kredi Politikaları ve Karar Destek\8_SPSS\99_Dashboard\01_Output\";
 string yol= pth+filePath;
	try
{
   File.Delete(pth+filePath); 
}
catch
{ 
Console.WriteLine("Log: "+yol+" adresi degisti");
	long time = DateTime.Now.Ticks;
	yol=pth+time.ToString()+"_"+filePath;
}



    StringBuilder fileContent = new StringBuilder();

        foreach (var col in dataTable.Columns) 
        {
            fileContent.Append(col.ToString() + "|");
        }

        fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);

        foreach (DataRow dr in dataTable.Rows) 
        {
            foreach (var column in dr.ItemArray) 
            {
                fileContent.Append("\"" + column.ToString() + "\"|");
            }

            fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);
        }

       System.IO.File.WriteAllText(yol, fileContent.ToString(),Encoding.Default); 
//	    Encoding.GetEncoding("Windows-1254")
//		Encoding.GetEncoding("iso-8859-9")
//		Encoding.GetEncoding(1254);

	   
 Console.WriteLine("Success: " +yol );
 } 

public static void mail(){ 
/*Microsoft.Office.Interop.Outlook.Application OutlookObject = new Microsoft.Office.Interop.Outlook.Application();
 
//Outlook programına gönderilmek üzere MailItem nesnesinin bir instance oluşturuyoruz
Microsoft.Office.Interop.Outlook.MailItem MailObject = (Microsoft.Office.Interop.Outlook.MailItem)(OutlookObject.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem));
 
//Mesajı gönderen "TO"
MailObject.To = "aozen";
//İhtiyaca göre "CC" ve "BCC" eklenmesi
//MailObject.CC = ccTextBox.Text;
//MailObject.BCC = bccTextBox.Text;
 
// Mail başlığının eklenmesi
MailObject.Subject = "Mail Başlığı";
 
MailObject.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
 
// Mesaj içeriği metin yada keyfinize göre HTML ekleyebilirsiniz
MailObject.HTMLBody = "Mail İçeriği";
 
// Mail'e attachment yani ek eklenmsi, birden çok ek te ekleyebilirsiniz
//MailObject.Attachments.Add("C:\\Users\\turhany\\Desktop\\a.png", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, "Ek Adı");
 
// Hazırladığınız mail template istediğiniz yere kaydetmenizi sağlar,
//ayrıca file dialog açar ve istediğiniz yeri seçmenize de izin verir
//MailObject.SaveAs(@"C:\demo.msg", Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
 
//Maili outlook açılarak içinde pencere olarak göterilemesini sağlar gönder butonuna siz basarsınız
//MailObject.Display();
 
//Maili direk yollamak isterseniz bu kodu kullanırsınız, pencere gösterilmeden direk yollanır. 
//(Bunu kullanacaksanız Display kodunu kapatmak gerek)
MailObject.Send();  
	*/
	
  /*   var oApp = new Outlook.Application();

        Microsoft.Office.Interop.Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
        var f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
      //  Thread.Sleep(5000); // a bit of startup grace time.
        var mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
        mailItem.Subject = "Error Report from user: " + AuthenticationManager.LoggedInUserName;
        mailItem.HTMLBody = "Test email\n"+ReadSignature();
        mailItem.To =  "aozen";
       // mailItem.Display(true);	
		mailItem.Send();
		*/
		
	}		
	
		
	/*	public static DataTable ConvertCSVtoDataTable(string strFilePath)
 {
            StreamReader sr = new StreamReader(strFilePath);
            string[] headers = sr.ReadLine().Split(','); 
            DataTable dt = new DataTable();
            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }
            return dt;
 } 
 */
   
 public static void createXLS(DataTable dt)        {
            FileStream stream = new FileStream("YAPILANDIRMALAR.xls", FileMode.OpenOrCreate);
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
