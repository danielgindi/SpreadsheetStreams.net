using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetStreams.Util
{
    public class WriteStreamWrapper : Stream
    {
        private Stream _Stream = null;
        private long _Position = 0;

        public WriteStreamWrapper(Stream stream)
        {
            _Stream = stream;
        }

        public override bool CanRead => _Stream.CanRead;
        public override bool CanSeek => _Stream.CanSeek;
        public override bool CanTimeout => _Stream.CanTimeout;
        public override bool CanWrite => _Stream.CanWrite;
        public override long Length => _Stream.Length;

        public override long Position
        {
            get { return _Position; }
            set { _Position = value; }
        }

        public override int ReadTimeout { get => _Stream.ReadTimeout; set => _Stream.ReadTimeout = value; }
        public override int WriteTimeout { get => _Stream.WriteTimeout; set => _Stream.WriteTimeout = value; }

        public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            return _Stream.BeginRead(buffer, offset, count, callback, state);
        }

        public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            return _Stream.BeginWrite(buffer, offset, count, ar => {
                _Position += count;
                if (callback != null)
                    callback.Invoke(ar);
            }, state);
        }

        public override void Close()
        {
            _Stream.Close();
        }

        public override Task CopyToAsync(Stream destination, int bufferSize, CancellationToken cancellationToken)
        {
            return _Stream.CopyToAsync(destination, bufferSize, cancellationToken);
        }

        public override int EndRead(IAsyncResult asyncResult)
        {
            var ret = _Stream.EndRead(asyncResult);
            _Position += ret;
            return ret;
        }

        public override void EndWrite(IAsyncResult asyncResult)
        {
            _Stream.EndWrite(asyncResult);
        }

        public override bool Equals(object obj)
        {
            return obj is WriteStreamWrapper && _Stream.Equals(((WriteStreamWrapper)obj)._Stream);
        }

        public override void Flush()
        {
            _Stream.Flush();
        }

        public override Task FlushAsync(CancellationToken cancellationToken)
        {
            return _Stream.FlushAsync(cancellationToken);
        }

        public override int GetHashCode()
        {
            return _Stream.GetHashCode();
        }

        public override object InitializeLifetimeService()
        {
            return _Stream.InitializeLifetimeService();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var ret = _Stream.Read(buffer, offset, count);
            _Position += ret;
            return ret;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return _Stream.ReadAsync(buffer, offset, count, cancellationToken)
                .ContinueWith(x => { _Position += x.Result; return x.Result; });
        }

        public override int ReadByte()
        {
            return _Stream.ReadByte();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            var ret = _Stream.Seek(offset, origin);
            _Position = _Stream.Position;
            return ret;
        }

        public override void SetLength(long value)
        {
            _Stream.SetLength(value);
        }

        public override string ToString()
        {
            return base.ToString();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _Stream.Write(buffer, offset, count);
            _Position += count;
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return _Stream.WriteAsync(buffer, offset, count, cancellationToken)
                .ContinueWith(x => _Position += count);
        }

        public override void WriteByte(byte value)
        {
            _Stream.WriteByte(value);
        }

        protected override void Dispose(bool disposing)
        {
            if (_Stream != null)
            {
                _Stream.Dispose();
            }
        }
    }
}
