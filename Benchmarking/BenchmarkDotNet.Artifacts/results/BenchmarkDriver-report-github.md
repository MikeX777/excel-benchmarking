```

BenchmarkDotNet v0.13.7, Windows 11 (10.0.22621.2134/22H2/2022Update/SunValley2)
Intel Core i9-9880H CPU 2.30GHz, 1 CPU, 16 logical and 8 physical cores
.NET SDK 6.0.100
  [Host]     : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT AVX2
  DefaultJob : .NET 6.0.0 (6.0.21.52210), X64 RyuJIT AVX2


```
|        Method | Mean | Error |
|-------------- |-----:|------:|
|      NpoiTest |   NA |    NA |
|    EPPlusTest |   NA |    NA |
| ClosedXmlTest |   NA |    NA |

Benchmarks with issues:
  BenchmarkDriver.NpoiTest: DefaultJob
  BenchmarkDriver.EPPlusTest: DefaultJob
  BenchmarkDriver.ClosedXmlTest: DefaultJob
