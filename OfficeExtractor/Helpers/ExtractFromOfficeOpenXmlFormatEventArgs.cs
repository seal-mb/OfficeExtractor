using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;

namespace OfficeExtractor.Helpers
{
    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Values that represent office XML document types. </summary>
    ///
    /// <remarks>   seal-mb, 19.01.2023. </remarks>
    ///-------------------------------------------------------------------------------------------------

    public enum OfficeXmlDocumentType : int
    {
        DocType_Unknown = -1,
        DocType_Word = 0,
        DocType_Excel = 1,
        DocType_PowerPoint = 2,
        DocType_Visio = 3,
    }

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Additional information for extract from office open XML format events. </summary>
    ///
    /// <remarks>   seal-mb, 19.01.2023. </remarks>
    ///-------------------------------------------------------------------------------------------------

    public class ExtractFromOfficeOpenXmlFormatEventArgs : EventArgs
    {
        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor. </summary>
        ///
        /// <remarks>   Martin.breer, 22.09.2016. </remarks>
        ///
        /// <param name="pack">     The package. </param>
        /// <param name="part">     The current part from the package. </param>
        /// <param name="doctype">  The XML document type. </param>
        ///-------------------------------------------------------------------------------------------------

        public ExtractFromOfficeOpenXmlFormatEventArgs(Package pack, PackagePart part, OfficeXmlDocumentType doctype, String embeddingPartString)
        {
            this.CurrentPackage = pack;
            this.CurrentPart = part;
            this.DocType = doctype;
            this.EmbeddingPartString = embeddingPartString;
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the current package. </summary>
        ///
        /// <value> The current package. </value>
        ///-------------------------------------------------------------------------------------------------

        public Package CurrentPackage { get; private set; }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the current part. </summary>
        ///
        /// <value> The current part. </value>
        ///-------------------------------------------------------------------------------------------------

        public PackagePart CurrentPart { get; private set; }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the type of the document. </summary>
        ///
        /// <value> The type of the document. </value>
        ///-------------------------------------------------------------------------------------------------

        public OfficeXmlDocumentType DocType { get; private set; }

        public String EmbeddingPartString { get; private set; }
    }

    

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Additional information for before extract from office open XML format events. </summary>
    ///
    /// <remarks>   Martin.breer, 22.09.2016. </remarks>
    ///-------------------------------------------------------------------------------------------------

    public class BeforeExtractFromOfficeOpenXmlFormatEventArgs : ExtractFromOfficeOpenXmlFormatEventArgs
    {
        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor. </summary>
        ///
        /// <remarks>   Martin.breer, 22.09.2016. </remarks>
        ///
        /// <param name="pack">     The package. </param>
        /// <param name="part">     The current part from the package. </param>
        /// <param name="doctype">  The XML document type. </param>
        ///-------------------------------------------------------------------------------------------------

        public BeforeExtractFromOfficeOpenXmlFormatEventArgs(Package pack, PackagePart part, OfficeXmlDocumentType doctype, String embeddingPartString)
            : base(pack, part, doctype, embeddingPartString)
        {

            this.Cancel = false;
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets a value indicating whether the cancel. </summary>
        ///
        /// <value> true if cancel, false if not. </value>
        ///-------------------------------------------------------------------------------------------------

        public bool Cancel { get; set; }
    }

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Additional information for after extract from office open XML format events. </summary>
    ///
    /// <remarks>   Martin.breer, 22.09.2016. </remarks>
    ///-------------------------------------------------------------------------------------------------

    public class AfterExtractFromOfficeOpenXmlFormatEventArgs : ExtractFromOfficeOpenXmlFormatEventArgs
    {
        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor. </summary>
        ///
        /// <remarks>   Martin.breer, 22.09.2016. </remarks>
        ///
        /// <param name="pack">     The package. </param>
        /// <param name="part">     The current part from the package. </param>
        /// <param name="doctype">  The XML document type. </param>
        /// <param name="filename"> Name of and path of the new extracted file. </param>
        ///-------------------------------------------------------------------------------------------------

        public AfterExtractFromOfficeOpenXmlFormatEventArgs(Package pack, PackagePart part, OfficeXmlDocumentType doctype, String embeddingPartString, String filename)
            : base(pack, part, doctype,embeddingPartString)
        {

            this.FileName = filename;

        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the filename of the extracted file. </summary>
        ///
        /// <value> The name of the file. </value>
        ///-------------------------------------------------------------------------------------------------

        public String FileName { get; private set; }
    }

    ///-------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Additional information for error extract from office open XML format events./
    /// </summary>
    ///-------------------------------------------------------------------------------------------------

    public class ErrorExtractFromOfficeOpenXmlFormatEventArgs : ExtractFromOfficeOpenXmlFormatEventArgs
    {
        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor./ </summary>
        ///
        /// <param name="pack">                 The pack. </param>
        /// <param name="part">                 The part. </param>
        /// <param name="doctype">              The doctype. </param>
        /// <param name="embeddingPartString">  The embedding part string. </param>
        /// <param name="ex">                   The ex. </param>
        ///-------------------------------------------------------------------------------------------------

        public ErrorExtractFromOfficeOpenXmlFormatEventArgs(Package pack, PackagePart part, OfficeXmlDocumentType doctype, String embeddingPartString, Exception ex)
            : base(pack, part, doctype, embeddingPartString)
        {
            this.LastException = ex;
            this.Cancel = false;
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets a value indicating whether the cancel./ </summary>
        ///
        /// <value> true if cancel, false if not. </value>
        ///-------------------------------------------------------------------------------------------------

        public bool Cancel { get;  set; }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the last exception./ </summary>
        ///
        /// <value> The last exception. </value>
        ///-------------------------------------------------------------------------------------------------

        public Exception LastException { get; private set; }
    }

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Delegate before office XML extract. </summary>
    ///
    /// <remarks>   Martin.breer, 22.09.2016. </remarks>
    ///
    /// <param name="currentExtractor"> The current extractor. </param>
    /// <param name="args">             Before extract from office open XML format event information. </param>
    ///-------------------------------------------------------------------------------------------------

    public delegate void DelegateBeforeOfficeXmlExtract(OfficeExtractor.Extractor currentExtractor, BeforeExtractFromOfficeOpenXmlFormatEventArgs args );

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Delegate after office XML exctract. </summary>
    ///
    /// <remarks>   Martin.breer, 22.09.2016. </remarks>
    ///
    /// <param name="currentExtractor"> The current extractor. </param>
    /// <param name="args">             After extract from office open XML format event information. </param>
    ///-------------------------------------------------------------------------------------------------

    public delegate void DelegateAfterOfficeXmlExctract(OfficeExtractor.Extractor currentExtractor, AfterExtractFromOfficeOpenXmlFormatEventArgs args);

    ///-------------------------------------------------------------------------------------------------
    /// <summary>   Delegate error office XML exctract. </summary>
    ///
    /// <param name="currentExtractor"> The current extractor. </param>
    /// <param name="args">             Error extract from office open XML format event information.
    /// </param>
    ///-------------------------------------------------------------------------------------------------

    public delegate void DelegateErrorOfficeXmlExctract(OfficeExtractor.Extractor currentExtractor, ErrorExtractFromOfficeOpenXmlFormatEventArgs args);

}
