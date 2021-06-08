using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Resources;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.ComponentModel;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Component;
using FS = NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Extractor;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;

namespace PBC.PipelineComponents.XLSX.to.Xml
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]

    [System.Runtime.InteropServices.Guid("88B5BF31-52B1-4D13-A436-E25492644D3B")]

    [ComponentCategory(CategoryTypes.CATID_Decoder)]

    public class XlsxtoCSV : Microsoft.BizTalk.Component.Interop.IComponent, IBaseComponent, IPersistPropertyBag, IComponentUI
    {
        private string fileName;
        public bool IsFirstRowHeader { get; set; }
        public string TempFolder { get; set; }


        #region IBaseComponent members

        /// <summary>
        /// Description of the component
        /// </summary>
        public string Description
        {
            get { return "Pipeline component to convert excel file to xml."; }
        }

        /// <summary>
        /// Name of the component
        /// </summary>
        [Browsable(false)]
        public string Name
        {
            get { return "Excel to xml converter"; }
        }

        /// <summary>
        /// Version of the component
        /// </summary>
        [Browsable(false)]
        public string Version
        {
            get { return "1.0.0.0"; }
        }

        #endregion

        #region IPersistPropertyBag members

        private string _namespace;
        private string _rootNodeName;
        private string _childNodeName;


        public string Namespace
        {
            get { return _namespace; }
            set { _namespace = null; }
        }

        public string RootNodeName
        {
            get { return _rootNodeName; }
            set { _rootNodeName = null; }
        }

        public string ChildNodeName
        {
            get { return _childNodeName; }
            set { _childNodeName = null; }
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            object val = null;
            val = ReadPropertyBag(propertyBag, "Namespace");
            if ((val != null))
                this._namespace = ((string)(val));

            val = ReadPropertyBag(propertyBag, "RootNodeName");
            if ((val != null))
                this._rootNodeName = ((string)(val));

            val = ReadPropertyBag(propertyBag, "ChildNodeName");
            if ((val != null))
                this._childNodeName = ((string)(val));
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            WritePropertyBag(propertyBag, "Namespace", this._namespace);
            WritePropertyBag(propertyBag, "RootNodeName", this._rootNodeName);
            WritePropertyBag(propertyBag, "ChildNodeName", this._childNodeName);
        }

        /// <summary>
        /// Gets class ID of component for usage from unmanaged code.
        /// </summary>
        /// <param name="classid">
        /// Class ID of the component
        /// </param>
        public void GetClassID(out System.Guid classid)
        {
            classid = new System.Guid("88B5BF31-52B1-4D13-A436-E25492644D3B");
        }

        /// <summary>
        /// not implemented
        /// </summary>
        public void InitNew()
        {

        }

        #region utility functionality

        /// <summary>
        /// Reads property value from property bag
        /// </summary>
        /// <param name="pb">Property bag</param>
        /// <param name="propName">Name of property</param>
        /// <returns>Value of the property</returns>
        private object ReadPropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName)
        {
            object val = null;
            try
            {
                pb.Read(propName, out val, 0);
            }
            catch (System.ArgumentException)
            {
                return val;
            }
            catch (System.Exception e)
            {
                throw new System.ApplicationException(e.Message);
            }
            return val;
        }



        /// <summary>
        /// Writes property values into a property bag.
        /// </summary>
        /// <param name="pb">Property bag.</param>
        /// <param name="propName">Name of property.</param>
        /// <param name="val">Value of property.</param>
        private void WritePropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName, object val)
        {
            try
            {
                pb.Write(propName, ref val);
            }
            catch (System.Exception e)
            {
                throw new System.ApplicationException(e.Message);
            }
        }

        #endregion

        #endregion

        #region IComponentUI members
        
        /// <summary>
        /// Component icon to use in BizTalk Editor
        /// </summary>
        [Browsable(false)]
        public IntPtr Icon
        {
            get
            {
                return new System.IntPtr();
            }
        }

        /// <summary>
        /// The Validate method is called by the BizTalk Editor during the build 
        /// of a BizTalk project.
        /// </summary>
        /// <param name="obj">An Object containing the configuration properties.</param>
        /// <returns>The IEnumerator enables the caller to enumerate through a collection of strings containing error messages. These error messages appear as compiler error messages. To report successful property validation, the method should return an empty enumerator.</returns>
        public System.Collections.IEnumerator Validate(object obj)
        {
            return null;
        }

        #endregion

        #region IComponent members

        /// <summary>
        /// Implements IComponent.Execute method.
        /// </summary>
        /// <param name="pc">Pipeline context</param>
        /// <param name="inmsg">Input message</param>
        /// <returns>Original input message</returns>
        /// <remarks>
        /// IComponent.Execute method is used to initiate
        /// the processing of the message in this pipeline component.
        /// </remarks>
        public Microsoft.BizTalk.Message.Interop.IBaseMessage Execute(Microsoft.BizTalk.Component.Interop.IPipelineContext context, Microsoft.BizTalk.Message.Interop.IBaseMessage inmsg)
        {
            Trace.WriteLine("Entering XLSX pipeline...");
            var tblData = new DataTable();
            var outMsg = context.GetMessageFactory().CreateMessage();
            var tempDir = string.Format("{0}\\{1}", TempFolder, Guid.NewGuid());

            try
            {
                if (inmsg == null || inmsg.BodyPart == null || inmsg.BodyPart.Data == null)
                {
                    throw new ArgumentNullException("pInMsg");
                }
                fileName = inmsg.Context.Read("ReceivedFileName", "http://schemas.microsoft.com/BizTalk/2003/file-properties").ToString();
                FileInfo fo = new FileInfo(fileName);
                fileName = fo.Name.ToString();
                Stream fs = inmsg.BodyPart.GetOriginalDataStream();
                XSSFWorkbook hssfworkbook = null;
                hssfworkbook = new XSSFWorkbook(fs);
                tblData = ConvertToDataTable(hssfworkbook);

                MemoryStream stream = WriteWorksheetToStream(tblData);

                stream.Seek(0, SeekOrigin.Begin);
                outMsg.AddPart("Body", context.GetMessageFactory().CreateMessagePart(), true);
                outMsg.BodyPart.Data = stream;

                fileName = "<string>" + fileName + "</string>";

                //Promote properties if required.
                for (int iProp = 0; iProp < inmsg.Context.CountProperties; iProp++)
                {
                    string strName;
                    string strNSpace = Namespace;
                    object val = inmsg.Context.ReadAt(iProp, out strName, out strNSpace);

                    // If the property has been promoted, respect the settings
                    if (inmsg.Context.IsPromoted(strName, strNSpace))
                        outMsg.Context.Promote(strName, strNSpace, val);
                    else
                        outMsg.Context.Write(strName, strNSpace, val);

                    //update the ReceivedFileName with the actual file entry name
                    if (strName == "ReceivedFileName")
                    {
                        outMsg.Context.Write(strName, strNSpace, fileName);
                    }
                }

                //To get Incoming message
                System.IO.Stream originalStream = outMsg.BodyPart.GetOriginalDataStream();

                //Working with XDocument
                XDocument xDoc;
                using (XmlReader reader = XmlReader.Create(originalStream))
                {
                    reader.MoveToContent();
                    xDoc = XDocument.Load(reader);
                }
                xDoc.Root.RemoveAttributes();
                xDoc.Root.Add(new XAttribute(XNamespace.Xmlns + "ns0", Namespace));

                //Added this piece to add namespace to all childnodes.
                var childNodes = xDoc.Descendants("Rate");
                foreach (var node in childNodes)
                {
                    node.Add(new XAttribute(XNamespace.Xmlns + "ns0", Namespace));
                    //XmlNode xmlNode = GetXmlNode(node);
                    //xmlNode.Prefix = "ns0";
                }


                //Added this piece to add prefix to all childnodes.
                XmlDocument doc = new XmlDocument();
                doc = GetXmlDocument(xDoc);

                foreach (XmlNode node in doc.SelectNodes("Rates/Rate"))
                {
                    if (node.Prefix.Length == 0)
                        node.Prefix = "ns0";
                }
                //doc.Save(@"C:\Users\KGannama\Desktop\test\xml\Test.xml");


                // Returning stream
                byte[] output = System.Text.Encoding.ASCII.GetBytes(xDoc.ToString());
                MemoryStream memoryStream = new MemoryStream();
                memoryStream.Write(output, 0, output.Length);
                memoryStream.Position = 0;
                outMsg.BodyPart.Data = memoryStream;
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
            }
            return outMsg;
        }

        private Stream GenerateStream(string s)
        {
            MemoryStream ms = new MemoryStream();
            StreamWriter sw = new StreamWriter(ms);
            sw.Write(s);
            sw.Flush();
            ms.Position = 0;
            return ms;
        }

        public MemoryStream WriteWorksheetToStream(DataTable data)
        {
            var ds = new DataSet("Rates");
            var stream = new MemoryStream();
            ds.Tables.Add(data);
            data.DataSet.WriteXml(stream);
            return stream;
        }

        private DataTable ConvertToDataTable(XSSFWorkbook hssfworkbook)
        {
            ISheet sheet = hssfworkbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            DataTable dt1 = new DataTable() { TableName = "Rate" };

            bool flag = true;
            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                if (flag)
                {
                    for (int i = 0; i < 70; i++)
                    {
                        ICell cell = row.GetCell(i);
                        if (cell == null)
                        {
                            dt1.Columns.Add("");
                        }
                        else
                        {
                            dt1.Columns.Add(cell.ToString().Trim());
                        }
                    }
                    flag = false;
                }
                else
                {
                    DataRow dr = dt1.NewRow();
                    for (int i = 0; i < 70; i++)
                    {
                        ICell cell = row.GetCell(i);
                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            dr[i] = cell.ToString().Trim();
                        }
                    }
                    dt1.Rows.Add(dr);
                }
            }
            return dt1;

            //System.IO.StringWriter writer = new System.IO.StringWriter();
            //dt1.WriteXml(@"C:\MyDataset.xml", true);
        }


        public static XmlDocument GetXmlDocument(XDocument document)
        {
            using (XmlReader xmlReader = document.CreateReader())
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlReader);
                if (document.Declaration != null)
                {
                    XmlDeclaration dec = xmlDoc.CreateXmlDeclaration(document.Declaration.Version,
                        document.Declaration.Encoding, document.Declaration.Standalone);
                    xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
                }
                return xmlDoc;
            }
        }

        public static XmlNode GetXmlNode(XElement element)
        {
            using (XmlReader xmlReader = element.CreateReader())
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlReader);
                return xmlDoc;
            }
        }

        public static XElement GetXElement(XmlNode node)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                node.WriteTo(xmlWriter);
            return xDoc.Root;
        }

        #endregion
    }
}







