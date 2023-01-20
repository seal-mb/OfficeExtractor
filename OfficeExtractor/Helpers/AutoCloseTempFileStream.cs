///-------------------------------------------------------------------------------------------------
// project:		OfficeExtractor
//
// file:		OfficeExtractor\Helpers\AutoCloseTempFileStream.cs
//
// author:			seal-mb  
// mail-to:     	martin.breer@sealsystems.de
// company name:	SEAL Systems
//
// copyright:		Copyright (c) 2023 SEAL Systems. All rights reserved.
//
// date:			20.01.2023
//
//--------------------------------------------------------------------
//
// summary:	Implements the automatic close temporary file stream class
//
///-------------------------------------------------------------------------------------------------

#region Using
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
#endregion

namespace OfficeExtractor.Helpers
{
    ///-------------------------------------------------------------------------------------------------
    /// <summary>   An automatic close temporary file stream. </summary>
    ///
    /// <remarks>   seal-mb, 20.01.2023. </remarks>
    ///-------------------------------------------------------------------------------------------------

    internal class AutoCloseTempFileStream : IDisposable
    {
        #region IDisposable

        private bool _disposedValue = false;    ///< True if disposed.

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged
        /// resources.
        /// </summary>
        ///
        /// <remarks>   Martin, 20.01.2023. </remarks>
        ///
        /// <param name="disposing">    True to release both managed and unmanaged resources; false to
        ///                             release only unmanaged resources. </param>
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        protected virtual void Dispose ( bool disposing )
        {
            if ( !_disposedValue )
            {

                _disposedValue = true;
                if ( disposing )
                {
                    // TODO: dispose managed state (managed objects)
                    _tempFile.Close ();
                    _tempFile.Dispose ();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalize
                // TODO: set large fields to null
                try
                {
                    if ( System.IO.File.Exists ( _fileName ) )
                    {
                        System.IO.File.Delete ( _fileName );
                    }
                }
                catch 
                {
                    
                }

            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>   Finalizer. </summary>
        ///
        /// <remarks>   Martin, 20.01.2023. </remarks>
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        ~AutoCloseTempFileStream ()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose ( disposing: false );
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged
        /// resources.
        /// </summary>
        ///
        /// <remarks>   Martin, 20.01.2023. </remarks>
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        public void Dispose ()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose ( disposing: true );
            GC.SuppressFinalize ( this );
        }

        #endregion

        #region Member


        private String _fileName;   ///< Filename of the file

        private Stream _tempFile;   ///< The temporary file

        #endregion

        #region Type Cast

        ///-------------------------------------------------------------------------------------------------
        /// <summary>
        /// Implicit cast that converts the given Stream to an AutoCloseTempFileStream.
        /// </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="stream">   The stream. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        ///-------------------------------------------------------------------------------------------------

        public static implicit operator AutoCloseTempFileStream ( Stream stream )
        {
            return new AutoCloseTempFileStream ( stream, true );
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>
        /// Implicit cast that converts the given byte[] to an AutoCloseTempFileStream.
        /// </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="data"> The data. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        ///-------------------------------------------------------------------------------------------------

        public static implicit operator AutoCloseTempFileStream ( byte[] data )
        {
            return new AutoCloseTempFileStream ( data );
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>
        /// Implicit cast that converts the given AutoCloseTempFileStream to a Stream.
        /// </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="stream">   The stream. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        ///-------------------------------------------------------------------------------------------------

        public static implicit operator Stream ( AutoCloseTempFileStream stream )
        {
            return stream.FileStream;
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>
        /// Implicit cast that converts the given AutoCloseTempFileStream to a byte[].
        /// </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="stream">   The stream. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        ///-------------------------------------------------------------------------------------------------

        public static implicit operator byte[] ( AutoCloseTempFileStream stream )
        {
            return stream.GetBytes ( true ).ToArray ();
        }

        #endregion

        #region Constructors

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Default constructor. </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///-------------------------------------------------------------------------------------------------

        public AutoCloseTempFileStream ()
        {
            _fileName = Path.GetTempFileName ();
            _tempFile = new FileStream ( _fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose );
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor. </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="inputStream">          Stream to read data from. </param>
        /// <param name="seekBegin">            (Optional) True to seek begin. </param>
        /// <param name="disposeInputStream">   (Optional) True to dispose input stream. </param>
        ///-------------------------------------------------------------------------------------------------

        public AutoCloseTempFileStream ( Stream inputStream, bool seekBegin = true, bool disposeInputStream = false )
            : this ()
        {
            try
            {
                if ( null != inputStream && inputStream.CanRead )
                {
                    long? posNow = null;
                    if ( seekBegin )
                    {
                        posNow = inputStream.Position;
                        inputStream.Position = 0;
                    }

                    inputStream.CopyTo ( _tempFile );
                    _tempFile.Seek ( 0, SeekOrigin.Begin );

                    if ( seekBegin && posNow.HasValue && !disposeInputStream )
                    {
                        inputStream.Position = posNow.Value;
                    }

                    if( disposeInputStream )
                    {
                        inputStream.Dispose ();
                    }
                }
                
            }
            catch 
            {
                throw;
            }
        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Constructor. </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="data"> Byte data to set. </param>
        ///-------------------------------------------------------------------------------------------------

        public AutoCloseTempFileStream ( byte[] data )
            : this ()
        {
            if ( data.Length > 0 )
            {
                _tempFile.Write ( data, 0, data.Length );
                _tempFile.Seek ( 0, SeekOrigin.Begin );
            }
        }

        #endregion

        #region Properties

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets the filename of the temporary file. </summary>
        ///
        /// <value> The name of the file. </value>
        ///-------------------------------------------------------------------------------------------------

        public String FileName => _fileName;

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets the file stream. </summary>
        ///
        /// <value> The file stream. </value>
        ///-------------------------------------------------------------------------------------------------

        public Stream FileStream => _tempFile;

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets the stream size. </summary>
        ///
        /// <value> The size of the stream. </value>
        ///-------------------------------------------------------------------------------------------------

        public long Length => _tempFile.Length;

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets or sets the position in file stream. </summary>
        ///
        /// <value> The position in the stream. </value>
        ///-------------------------------------------------------------------------------------------------
        public long Position
        {
            get => _tempFile.Position;
            set => _tempFile.Position = value;
        }

        #endregion

        #region Methodes


        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Gets the bytes in the stream. </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <param name="savePos">  (Optional) True to save position. </param>
        ///
        /// <returns>
        /// Get the bytes from the stream.
        /// </returns>
        ///-------------------------------------------------------------------------------------------------
        public IEnumerable<byte> GetBytes ( bool savePos = true )
        {
            var position = _tempFile.Position;


            for ( int item = _tempFile.ReadByte (); item != -1; item = _tempFile.ReadByte () )
            {
                yield return ( byte )item;
            }

            if ( savePos )
            {
                _tempFile.Position = position;
            }

        }

        ///-------------------------------------------------------------------------------------------------
        /// <summary>   Reads the byte from stream. </summary>
        ///
        /// <remarks>   seal-mb, 20.01.2023. </remarks>
        ///
        /// <returns>   The byte. </returns>
        ///-------------------------------------------------------------------------------------------------

        public int ReadByte()
        { 
            return _tempFile.ReadByte (); 
        }

        #endregion


    }
}
