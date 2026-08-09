using BenchmarkDotNet.Running;

// dotnet run -c Release --project benchmarks/DocToolkit.Benchmarks
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
