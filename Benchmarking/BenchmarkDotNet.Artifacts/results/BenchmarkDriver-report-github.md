```

BenchmarkDotNet v0.13.7, Windows 11 (10.0.22621.2134/22H2/2022Update/SunValley2)
12th Gen Intel Core i7-12700H, 1 CPU, 20 logical and 14 physical cores
.NET SDK 7.0.203
  [Host]     : .NET 6.0.15 (6.0.1523.11507), X64 RyuJIT AVX2
  DefaultJob : .NET 6.0.15 (6.0.1523.11507), X64 RyuJIT AVX2


```
|          Method |           Mean |       Error |      StdDev |
|---------------- |---------------:|------------:|------------:|
|     TestForLoop |      0.0307 ns |   0.0117 ns |   0.0104 ns |
| TestLinqForEach | 70,937.3454 ns | 766.4127 ns | 716.9029 ns |
|     TestForEach | 22,665.1254 ns | 168.6759 ns | 157.7796 ns |
