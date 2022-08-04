using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Nerosoft.Powersheet;
using Nerosoft.Powersheet.Npoi;

namespace Powersheet.Npoi.Test;

[SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
[SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "<Pending>")]
public class Startup
{
    public void ConfigureHost(IHostBuilder hostBuilder)
    {
        hostBuilder.ConfigureAppConfiguration(builder =>
                   {
                       builder.AddJsonFile("appsettings.json");
                   })
                   .ConfigureServices((_, services) =>
                   {
                       services.AddSingleton<ISheetWrapper, SheetWrapper>();
                   });
    }

    // ConfigureServices(IServiceCollection services)
    // ConfigureServices(IServiceCollection services, HostBuilderContext hostBuilderContext)
    // ConfigureServices(HostBuilderContext hostBuilderContext, IServiceCollection services)
    public void ConfigureServices(IServiceCollection services, HostBuilderContext hostBuilderContext)
    {
    }

    //public void Configure(IServiceProvider applicationServices, IIdGenerator idGenerator)
    //{
    //  InitData();
    //}

    public void Configure(IServiceProvider applicationServices)
    {
        //var config = applicationServices.GetService<IConfiguration>();
    }
}