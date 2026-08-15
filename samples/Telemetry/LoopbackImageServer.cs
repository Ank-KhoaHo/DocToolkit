using System.Net;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// A tiny HTTP server bound to <c>127.0.0.1</c>, owned and torn down entirely by this sample.
///
/// The library opens no socket by default, and this sample does not weaken that: it never reaches
/// the real internet, and nothing here answers on any address a machine's network interface is
/// actually listening on. It exists only so the "ok" telemetry outcome - a fetch that actually
/// succeeds - has something to succeed against, without this sample depending on anything outside
/// its own process.
///
/// It answers every request with the same small PNG - there is exactly one thing to serve - and
/// counts accepted connections so the console output can show, not just claim, how many times
/// <c>GuardedResourceLoader</c> actually reached out.
/// </summary>
internal sealed class LoopbackImageServer : IDisposable
{
    private readonly TcpListener _listener;
    private readonly CancellationTokenSource _stopping = new();
    private readonly byte[] _image;
    private int _connections;

    public LoopbackImageServer(byte[] image)
    {
        _image = image;
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _ = AcceptLoopAsync();
    }

    public int Port { get; }

    public string BaseUrl => $"http://127.0.0.1:{Port}";

    /// <summary>How many TCP connections this server has accepted since it started.</summary>
    public int Connections => Volatile.Read(ref _connections);

    private async Task AcceptLoopAsync()
    {
        while (!_stopping.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await _listener.AcceptTcpClientAsync(_stopping.Token);
            }
            catch (Exception)
            {
                return; // Stopped, or being disposed - nothing left to accept.
            }

            Interlocked.Increment(ref _connections);
            _ = ServeAsync(client);
        }
    }

    private async Task ServeAsync(TcpClient client)
    {
        using (client)
        {
            try
            {
                var stream = client.GetStream();

                // The request is tiny and arrives in one segment on loopback. This sample does not
                // need to parse it - only to drain it before writing a response, the same courtesy
                // a real server extends a client that has not finished sending yet.
                var buffer = new byte[2048];
                await stream.ReadAsync(buffer, _stopping.Token);

                var header = Encoding.ASCII.GetBytes(
                    "HTTP/1.1 200 OK\r\n" +
                    "Content-Type: image/png\r\n" +
                    $"Content-Length: {_image.Length}\r\n" +
                    "Connection: close\r\n\r\n");

                await stream.WriteAsync(header, _stopping.Token);
                await stream.WriteAsync(_image, _stopping.Token);
                await stream.FlushAsync(_stopping.Token);
            }
            catch (Exception)
            {
                // The caller hung up, or the server is shutting down. Either way there is nothing
                // left to do - the connection was already counted, which is this class's one job.
            }
        }
    }

    public void Dispose()
    {
        _stopping.Cancel();
        _listener.Stop();
        _stopping.Dispose();
    }
}
