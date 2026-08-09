using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The smallest <see cref="IOptionsMonitor{T}"/> that lets a unit test construct a service
/// directly, replacing the <c>Options.Create</c> these tests used before the services moved from
/// <c>IOptions</c> to <c>IOptionsMonitor</c>.
///
/// Setting <see cref="CurrentValue"/> is enough because the services read it on every call rather
/// than caching it — which is the property under test. <see cref="OnChange"/> is unused and returns
/// null: nothing here subscribes, and a change token nobody reads would be scaffolding pretending
/// to be behaviour.
///
/// The end-to-end reload behaviour is proven through the real container in
/// <c>ServiceCollectionExtensionsTests</c>, by socket count. This type exists for the unit tests
/// that never involved a container in the first place.
/// </summary>
internal sealed class TestOptionsMonitor<T> : IOptionsMonitor<T>
{
    public TestOptionsMonitor(T value) => CurrentValue = value;

    public T CurrentValue { get; set; }

    public T Get(string? name) => CurrentValue;

    public IDisposable? OnChange(Action<T, string?> listener) => null;
}
