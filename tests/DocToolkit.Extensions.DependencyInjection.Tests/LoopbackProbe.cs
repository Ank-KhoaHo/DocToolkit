using System.Net;
using System.Net.Sockets;
using System.Text;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// A minimal loopback TCP listener, proving whether <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// registered through <see cref="ServiceCollectionExtensions.AddDocToolkit"/> actually reached the
/// converter. Answers every connection with a tiny valid image so a fetch completes cleanly
/// instead of hanging or erroring - the assertion only cares whether a connection was accepted.
/// </summary>
internal sealed class LoopbackProbe : IDisposable
{
    private readonly TcpListener _listener;
    private readonly CancellationTokenSource _stopping = new();
    private int _connections;

    public LoopbackProbe()
    {
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _ = AcceptLoopAsync();
    }

    public int Port { get; }

    public string ImageUrl => $"http://127.0.0.1:{Port}/x.gif";

    public int Connections => Volatile.Read(ref _connections);

    public async Task<bool> WaitForConnectionAsync(TimeSpan timeout)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        while (stopwatch.Elapsed < timeout)
        {
            if (Connections > 0) return true;
            await Task.Delay(25);
        }

        return Connections > 0;
    }

    private async Task AcceptLoopAsync()
    {
        while (!_stopping.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await _listener.AcceptTcpClientAsync(_stopping.Token);
            }
            catch
            {
                return; // Listener stopped, or the test finished.
            }

            Interlocked.Increment(ref _connections);
            _ = RespondAsync(client);
        }
    }

    private static async Task RespondAsync(TcpClient client)
    {
        using (client)
        {
            try
            {
                var stream = client.GetStream();
                using var reader = new StreamReader(stream, leaveOpen: true);
                while (!string.IsNullOrEmpty(await reader.ReadLineAsync())) { }

                var body = System.Convert.FromHexString(
                    "47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b");
                var header = Encoding.ASCII.GetBytes(
                    $"HTTP/1.1 200 OK\r\nContent-Type: image/gif\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");

                await stream.WriteAsync(header);
                await stream.WriteAsync(body);
            }
            catch
            {
                // Best effort - the assertion only needs Connections to have been incremented.
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
