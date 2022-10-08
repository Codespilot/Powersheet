## Powersheet简介
Powersheet是一个基于.net standard 2.1开发的跨平台的Excel数据导入、导出类库。配置灵活，使用方便。支持DataTable、List<T>作为数据源。支持侵入和非侵入两种字段映射配置方式。

### 功能
1. 从Excel文件或流提取数据并转换为DataTable或对象集合；
2. 将DataTable或对象集合的数据写入到Excel文件；

### 特色
1. 基于.net standard 2.1开发，支持跨平台使用；
2. 轻量，不需要安装 Microsoft Office、COM+组件，体积小；
3. 提供nuget包，安装方便快捷；
4. 使用简单，对象属性和Excel列映射配置灵活，支持侵入式（使用属性Attribute）和非侵入（传入配置参数）两种方式；
5. 易于扩展；

## Nuget包
|包名称|版本|描述|
|--|--|--|
|Powersheet.Core|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Core.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Core/)|Powersheet核心类库，包含接口定义，公共方法实现以及其他需要用到的类。|
|Powersheet.Npoi|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Npoi.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Npoi/)|Powersheet的NPOI实现|
|Powersheet.Epplus|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Epplus.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Epplus/)|Powersheet的EPPLUS实现|

## 如何使用

### 安装Powersheet

> Powersheet.Epplus或Powersheet.Npoi根据需要安装其一即可。如果需要自行实现表格读取，仅需要引入Powersheet.Core。

命令行
```shell
dotnet add package Powersheet.Core 

dotnet add package Powersheet.Npoi 

dotnet add package Powersheet.Epplus 
```

Script & Interactive
```shell
#r "nuget: Powersheet.Core"

#r "nuget: Powersheet.Npoi"

#r "nuget: Powersheet.Epplus"
```

### 配置依赖注入

使用Microsoft.Extensions.DependencyInjection
```csharp
services.AddSingleton<ISheetWrapper, SheetWrapper>();
```

使用Autofac
```csharp
builder.RegisterType<SheetWrapper>().As<ISheetWrapper>().SingleInstance();
```

如果使用EPPLUS版本的实现，还需要在appSettings.json文件增加一个配置节点
```json
"EPPlus": {
    "ExcelPackage": {
      "LicenseContext": "NonCommercial"
    }
  }
```
### 接口定义

#### 读取表格内容到DataTable

```csharp
Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. options - 配置选项
> 3. sheetIndex - 表格位置索引，起始值为0

```csharp
Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. options - 配置选项
> 3. sheetName - 表格名称，不指定取第一个
```csharp
Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. options - 配置选项
> 3. sheetIndex - 表格位置索引，起始值为0
```csharp
Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. options - 配置选项
> 3. sheetName - 表格名称，不指定取第一个

#### 读取表格内容到对象集合
```csharp
Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. options - 配置选项
> 3. sheetIndex - 表格位置索引，起始值为0
> 4. T - 结果对象类型，必须是类且包含公开的无参构造器

```csharp
Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. options - 配置选项
> 3. sheetName - 表格名称，不指定取第一个
> 4. T - 结果对象类型，必须是类且包含公开的无参构造器

```csharp
Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. options - 配置选项
> 3. sheetIndex - 表格位置索引，起始值为0
> 4. T - 结果对象类型，必须是类且包含公开的无参构造器

```csharp
Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. options - 配置选项
> 3. sheetName - 表格名称，不指定取第一个
> 4. T - 结果对象类型，必须是类且包含公开的无参构造器

#### 读取指定单列数据到集合
```csharp
Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. firstRowNumber - 起始行号，从1开始
> 3. columnNumber - 列号，起始值为1
> 4. sheetIndex - 表格位置索引，起始值为0
> 5. valueConvert - 值转换方法

```csharp
Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert, CancellationToken cancellationToken);
```
> 1. file - 文件路径
> 2. firstRowNumber - 起始行号，起始值为1
> 3. columnNumber - 列号，起始值为1
> 4. sheetName - 表格名称，不指定取第一个
> 5. valueConvert - 值转换方法

```csharp
Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object,  CultureInfo, T> valueConvert, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. firstRowNumber - 起始行号，起始值为1
> 3. columnNumber - 列号，起始值为1
> 4. sheetIndex - 表格位置索引，起始值为0
> 5. valueConvert - 值转换方法

```csharp
Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert, CancellationToken cancellationToken);
```
> 1. stream - 文件流
> 2. firstRowNumber - 起始行号，起始值为1
> 3. columnNumber - 列号，起始值为1
> 4. sheetName - 表格名称，不指定取第一个
> 5. valueConvert - 值转换方法