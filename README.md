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
|包名称|版本|发布日期|
|--|--|--|
|Powersheet.Core|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Core.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Core/)|2021.07.20|
|Powersheet.Npoi|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Npoi.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Npoi/)|2021.07.20|
|Powersheet.Epplus|[![NuGet](https://img.shields.io/nuget/v/Powersheet.Epplus.svg?label=nuget)](https://www.nuget.org/packages/Powersheet.Epplus/)|2021.07.20|

## 如何使用

### 安装Powersheet

> Powersheet.Epplus或Powersheet.Npoi根据需要安装其一即可。

命令行
```ps
dotnet add package Powersheet.Core 

dotnet add package Powersheet.Npoi 

dotnet add package Powersheet.Epplus 
```

项目文件
```xml
<PackageReference Include="Powersheet.Core" Version="2021.7.20" />
<PackageReference Include="Powersheet.Npoi" Version="2021.7.20" />
<PackageReference Include="Powersheet.Epplus" Version="2021.7.20" />
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

如果使用EPPLUS版本的实现，还需要在appsetting.json文件增加一个配置节点
```json
"EPPlus": {
    "ExcelPackage": {
      "LicenseContext": "NonCommercial"
    }
  }
```
### 接口
