using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// A real TCP listener on 127.0.0.1 that counts accepted connections and records the request line
/// of each one.
///
/// Counting at the socket, not at an injected HTTP abstraction, is the point: it holds regardless
/// of which dependency does the fetching, whether it uses HttpClient, WebRequest or a raw socket,
/// and whether a future version of that dependency changes its mind. The connection is counted the
/// instant it is accepted - before a single byte is read - so a caller that connects and then gives
/// up is still caught.
///
/// Shared between <c>AirGapGuardTests</c>, which only needs the default responder to prove
/// silence, and <c>GuardedResourceLoaderTests</c>, which supplies its own <paramref name="respond"/>
/// delegate (via the constructor) to exercise <c>FetchAsync</c>'s byte cap, timeout and cancellation
/// behaviour against a real socket rather than a mock.
/// </summary>
internal sealed class LoopbackProbe : IDisposable
{
    /// <summary>How long to keep listening after a call returns before declaring "nothing connected".</summary>
    private static readonly TimeSpan SettleWindow = TimeSpan.FromMilliseconds(750);

    private readonly TcpListener _listener;
    private readonly CancellationTokenSource _stopping = new();
    private readonly ConcurrentQueue<string> _requestLines = new();
    private readonly ITestOutputHelper _output;
    private readonly Func<NetworkStream, string, CancellationToken, Task> _respond;
    private int _connections;

    /// <param name="output">Where to write diagnostics.</param>
    /// <param name="respond">
    /// How to answer an accepted connection, given the stream, the request line and a token that
    /// fires when the probe is disposed. Defaults to a plausible response: a 1x1 bitmap for
    /// anything image-shaped, a stylesheet otherwise - so the opt-in path has something real to
    /// embed and a leak on the default path would be a *successful* fetch rather than a connection
    /// error that might be swallowed silently.
    /// </param>
    /// <summary>
    /// Starts a loopback listener on an OS-assigned port and begins accepting connections.
    /// </summary>
    public LoopbackProbe(
        ITestOutputHelper output, Func<NetworkStream, string, CancellationToken, Task>? respond = null)
    {
        _output = output;
        _respond = respond ?? DefaultResponder;
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _ = AcceptLoopAsync();
    }

    public int Port { get; }

    public string BaseUrl => $"http://127.0.0.1:{Port}";

    public int Connections => Volatile.Read(ref _connections);

    /// <summary>Waits out the settle window, then requires that nothing ever connected.</summary>
    public async Task AssertSilentAsync(string api)
    {
        await Task.Delay(SettleWindow);

        _output.WriteLine(
            $"{api}: {Connections} connection(s) accepted on {BaseUrl} " +
            $"(listened a further {SettleWindow.TotalMilliseconds:0} ms after the call returned).");

        Xunit.Assert.True(Connections == 0,
            $"{api} opened {Connections} outbound connection(s) to {BaseUrl}" +
            (_requestLines.IsEmpty ? "." : $": {string.Join(" | ", _requestLines)}.") +
            " DocToolkit's users have no internet access, so this hangs and then fails on " +
            "their machines.");
    }

    public async Task<bool> WaitForConnectionAsync(TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < timeout)
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
            catch (Exception)
            {
                return; // Listener stopped, or the test finished. Nothing left to count.
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
                var requestLine = await ReadRequestLineAsync(stream);
                _requestLines.Enqueue(requestLine);

                await _respond(stream, requestLine, _stopping.Token);
            }
            catch (Exception)
            {
                // The caller hung up, or the probe is being torn down. The connection has
                // already been counted, which is all this class exists to do.
            }
        }
    }

    private static async Task DefaultResponder(NetworkStream stream, string requestLine, CancellationToken ct)
    {
        var css = requestLine.Contains(".css", StringComparison.OrdinalIgnoreCase);
        var body = css ? Encoding.ASCII.GetBytes("p { color: #333; }") : ImageFixtures.Bmp();

        await stream.WriteAsync(Encoding.ASCII.GetBytes(
            "HTTP/1.1 200 OK\r\n" +
            $"Content-Type: {(css ? "text/css" : "image/bmp")}\r\n" +
            $"Content-Length: {body.Length}\r\nConnection: close\r\n\r\n"), ct);
        await stream.WriteAsync(body, ct);
        await stream.FlushAsync(ct);
    }

    private static async Task<string> ReadRequestLineAsync(NetworkStream stream)
    {
        var buffer = new byte[2048];
        var seen = new StringBuilder();

        while (seen.ToString().IndexOf("\r\n", StringComparison.Ordinal) < 0 && seen.Length < 8192)
        {
            var read = await stream.ReadAsync(buffer);
            if (read == 0) break;
            seen.Append(Encoding.ASCII.GetString(buffer, 0, read));
        }

        var text = seen.ToString();
        var end = text.IndexOf("\r\n", StringComparison.Ordinal);
        return end < 0 ? text : text[..end];
    }

    public void Dispose()
    {
        _stopping.Cancel();
        _listener.Stop();
        _stopping.Dispose();
    }
}
