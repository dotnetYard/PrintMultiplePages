using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using PrintMultiplePages.ReportingService;

namespace PrintMultiplePages
{
    class PrintProcecssor
    {
        #region Prop
        public byte[][] RenderedReport
        {
            get
            {
                return m_renderedReport;
            }
            set
            {
                m_renderedReport = value;
            }
        }
        #endregion

        #region Private Field
        ReportExecutionService rs;
        private byte[][] m_renderedReport;
        private Graphics.EnumerateMetafileProc m_delegate = null;
        private MemoryStream m_currentPageStream;
        private Metafile m_metafile = null;
        private int m_numberOfPages, m_currentPrintingPage, m_lastPrintingPage;
        private bool Landscape;
        private string PageSizeOverride;

        #endregion

        #region Ctor
        public PrintProcecssor()
        {
            // Create proxy object and authenticate
            rs = new ReportExecutionService();
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        #endregion

        #region Public Method        

        public bool PrintReport(string printerName, string reportName, Int32 paperSourceIndex, string reportParamaters, bool landscape, string pageSizeOverride)
        {
            bool result = false;
            
            PageSizeOverride = pageSizeOverride;
            Landscape = landscape;
            RenderedReport = RenderReport(reportName, reportParamaters);

            PrinterSettings printerSettings = new PrinterSettings();
            try
            {                
                if (m_numberOfPages < 1)
                    return result;

                printerSettings.PrintRange = PrintRange.AllPages;
                printerSettings.PrinterName = printerName;
                PaperSource paperSource = printerSettings.PaperSources[paperSourceIndex];
                using (PrintDocument pd = new PrintDocument())
                {
                    m_currentPrintingPage = 1;
                    m_lastPrintingPage = m_numberOfPages;
                    pd.PrinterSettings = printerSettings;
                    pd.DefaultPageSettings.Landscape = landscape;

                    pd.DefaultPageSettings.PaperSource = paperSource;
                    if (PageSizeOverride == "Force 5.5 X 8.5")
                    {
                        pd.DefaultPageSettings.PaperSize = printerSettings.PaperSizes[6];
                    }

                    pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
                    pd.PrintController = new StandardPrintController();
                    pd.OriginAtMargins = true;
                    pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                    pd.Print();
                    pd.PrintPage -= new PrintPageEventHandler(pd_PrintPage);

                    result = true;
                }

            }
            finally
            {
                printerSettings = null;
            }

            return result;
        }
        #endregion

        #region Private Method 

        private byte[][] RenderReport(string reportPath, string reportParamaters)
        {
            // Private variables for rendering            
            string encoding, extension, mimeType, format = "IMAGE";
            string deviceInfo = String.Format(@"<DeviceInfo><OutputFormat>{0}</OutputFormat></DeviceInfo>", "emf");
            string[] streamIDs = null;
            Warning[] warnings = null;
            Byte[] firstPage = null;
            Byte[] nextPage = null;
            Byte[][] pages = null;

            ParameterValue[] parameters = null;
            if (reportParamaters != "")
            {
                char[] c = new char[1];
                c[0] = '|';
                char[] c2 = new char[1];
                c2[0] = '~';
                string[] stringParameters = reportParamaters.Split(c);
                int NumberOfParameters = stringParameters.Length;
                parameters = new ParameterValue[NumberOfParameters];
                for (int i = 0; i < NumberOfParameters; i++)
                {
                    parameters[i] = new ParameterValue();
                    parameters[i].Name = stringParameters[i].Split(c2)[0];
                    parameters[i].Value = stringParameters[i].Split(c2)[1];
                }
            }

            ExecutionHeader execHeader = new ExecutionHeader();
            rs.ExecutionHeaderValue = execHeader;

            ExecutionInfo execInfo = new ExecutionInfo();
            execInfo = rs.LoadReport(reportPath, null);

            rs.SetExecutionParameters(parameters, "en-us");
            String SessionId = rs.ExecutionHeaderValue.ExecutionID;

            //Exectute the report and get page count.           
            // Renders the first page of the report and returns streamIDs for 
            // subsequent pages
            firstPage = rs.Render(
                format,
                deviceInfo,
                out extension,
                out mimeType,
                out encoding,
                out warnings,
                out streamIDs);

            // The total number of pages of the report is 1 + the streamIDs  
            if (firstPage.Length > 0)
            {
                m_numberOfPages = 1;
                pages = new Byte[m_numberOfPages][];

                // The first page was already rendered
                pages[0] = firstPage;

                int pageIndex = m_numberOfPages;
                do
                {
                    deviceInfo =
                       String.Format(@"<DeviceInfo><OutputFormat>{0}</OutputFormat><StartPage>{1}</StartPage></DeviceInfo>",
                           "emf", ++pageIndex);
                    nextPage = rs.Render(format, deviceInfo, out extension, out mimeType, out encoding, out warnings, out streamIDs);
                    if (nextPage.Length > 0)
                    {
                        Array.Resize(ref pages, ++m_numberOfPages);
                        pages[m_numberOfPages - 1] = nextPage;
                    }
                }
                while (nextPage.Length > 0);
            }

            return pages;
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            ev.HasMorePages = false;
            if (m_currentPrintingPage <= m_lastPrintingPage && MoveToPage(m_currentPrintingPage))
            {
                // Draw the page
                ReportDrawPage(ev.Graphics);
                // If the next page is less than or equal to the last page, 
                // print another page.
                if (++m_currentPrintingPage <= m_lastPrintingPage)
                    ev.HasMorePages = true;
            }
        }

        // Method to draw the current emf memory stream 
        private void ReportDrawPage(Graphics g)
        {
            if (null == m_currentPageStream || 0 == m_currentPageStream.Length || null == m_metafile)
                return;
            lock (this)
            {
                // Set the metafile delegate.
                int width = m_metafile.Width;
                int height = m_metafile.Height;
                m_delegate = new Graphics.EnumerateMetafileProc(MetafileCallback);
                // Draw in the rectangle
                if (PageSizeOverride == "Original")
                {
                    Point destPoint = new Point(0, 0);
                    g.EnumerateMetafile(m_metafile, destPoint, m_delegate);
                }
                else if (PageSizeOverride == "Force 5.5 X 8.5")
                {
                    Point[] p = new Point[3];
                    p[0] = new Point(0, 0);
                    if (Landscape)
                    {
                        p[1] = new Point(859, 0);
                        p[2] = new Point(0, 568);
                    }
                    else
                    {
                        p[1] = new Point(568, 0);
                        p[2] = new Point(0, 859);
                    }
                    g.EnumerateMetafile(m_metafile, p, m_delegate);
                }
                else //Force 8.5 X 11
                {
                    Point[] p = new Point[3];
                    p[0] = new Point(0, 0);
                    if (Landscape)
                    {
                        p[1] = new Point(1118, 0);
                        p[2] = new Point(0, 859);
                    }
                    else
                    {
                        p[1] = new Point(859, 0);
                        p[2] = new Point(0, 1118);
                    }
                    g.EnumerateMetafile(m_metafile, p, m_delegate);
                }
                // Clean up
                m_delegate = null;
            }
        }

        private bool MoveToPage(Int32 page)
        {
            // Check to make sure that the current page exists in
            // the array list
            if (null == this.RenderedReport[m_currentPrintingPage - 1])
                return false;
            // Set current page stream equal to the rendered page
            m_currentPageStream = new MemoryStream(this.RenderedReport[m_currentPrintingPage - 1]);
            // Set its postion to start.
            m_currentPageStream.Position = 0;
            // Initialize the metafile
            if (null != m_metafile)
            {
                m_metafile.Dispose();
                m_metafile = null;
            }
            // Load the metafile image for this page
            m_metafile = new Metafile((Stream)m_currentPageStream);
            return true;
        }

        private bool MetafileCallback(EmfPlusRecordType recordType, int flags, int dataSize, IntPtr data, PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            // Dance around unmanaged code.
            if (data != IntPtr.Zero)
            {
                // Copy the unmanaged record to a managed byte buffer 
                // that can be used by PlayRecord.
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);
            }
            // play the record.      
            m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        #endregion        
    }
}
