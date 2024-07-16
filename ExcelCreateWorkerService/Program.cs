using ExcelCreateWorkerService;
using ExcelCreateWorkerService.Models;
using ExcelCreateWorkerService.Services;
using Microsoft.EntityFrameworkCore;
using RabbitMQ.Client;

var builder = Host.CreateApplicationBuilder(args);

string connectionString = builder.Configuration.GetConnectionString("AdventureWorks");
builder.Services.AddDbContext<AdventureWorks2022Context>(options =>
{
    options.UseSqlServer(connectionString);
});

string uriStr = builder.Configuration.GetConnectionString("RabbitMQ");
builder.Services.AddSingleton(serviceProvider => new ConnectionFactory()
{
    Uri = new Uri(uriStr),
    DispatchConsumersAsync = true
});

builder.Services.AddSingleton<RabbitMQClientService>();

builder.Services.AddHostedService<Worker>();

var host = builder.Build();
host.Run();
