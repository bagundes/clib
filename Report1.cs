using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace clib
{
    public class Report1 : IDisposable
    {
        private FileInfo ReportFile;
        private ReportDocument ReportDoc;
        public Report1(byte[] report, string fileName)
        {
            ReportFile = klib.Shell.CreateFile(report, fileName);

            ReportDoc = new ReportDocument();
            ReportDoc.Load(ReportFile.FullName); 
        }

        public Report1(string fileName)
        {
            ReportDoc = new ReportDocument();
            ReportDoc.Load(ReportFile.FullName);
        }

        public void Conection(string connString)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connString);
            ReportDoc.DataSourceConnections[0].SetConnection(builder.DataSource, 
                builder.InitialCatalog,
                builder.UserID,
                builder.Password);
        }

        public void Parameters(List<klib.model.Select> parameters)
        {
            foreach (var param in parameters)            
                ReportDoc.SetParameterValue(param.Name, param.Value.Value);            
        }

        public FileInfo Save(string fileName)
        {
            ReportDoc.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, fileName);
            return new FileInfo(fileName);
        }

        public void Print(klib.model.Printer printer)
        {
            ReportDoc.PrintOptions.PrinterName = printer.Name;
            ReportDoc.PrintToPrinter(1, false, 0, 0);
        }

        public void Dispose()
        {
            ReportFile = null;
            if (ReportDoc != null)
            {
                ReportDoc.Close();
                ReportDoc.Dispose();
                ReportDoc = null;
            }
        }
    }
}
