using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FaultlessExecution;

namespace ExcelToSqlServer
{
    internal class Program
    {
        internal static IServiceProvider ServiceProvider { get; set; }
        internal static IConfigurationRoot Configuration { get; set; }
        static void Main(string[] args)
        {
            string? fileName = args?.FirstOrDefault();

            //first and foremost, did they give us a file to process.  if not, they are terrible people and should never be allowed in public.
            if (string.IsNullOrEmpty(fileName))
            {
                Console.WriteLine("Bro, where's your file?");
                fileName = Console.ReadLine();
            }

            //fileexists does not like the file name to be in quotes, so lets trim whitespace and quotes
            fileName = fileName
                ?.Trim()
                ?.TrimEnd('"')
                ?.TrimStart('"');

            Console.WriteLine(fileName);

            if (System.IO.File.Exists(fileName))
            {

                IHost host;

                host = Startup();
                if (host != null)
                {
                    Run(fileName);
                }

                //   we must flush, otherwise the console might close out before logs are done writing
                Log.CloseAndFlush();
            }
            else
            {
                Console.WriteLine("Com'on man, that file doesn't even exist");
            }

            bool stayOpen = Configuration?.GetValue<bool>("StayOpen") ?? false;//it's possible bad things happened and Configuration never go initialized, therefor null check and default false because better to make someone review the log file than have a background process hungup, waiting for input

            if (stayOpen)
            {
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
            }

            return;
        }

        static IHost Startup()
        {

            IHost host;//this fella will get built in a try catch so we can track errors with intantiation versus errors running imports

            //do this first because serilogger wants to use it to get it's settings
            var configuration = CreateConfigurationRoot();
            Program.Configuration = configuration;

            //create an instance of serilogger, using appsettings, so we can log what happens on startup
            //NOTE:  this is just creating one instance, it is NOT setting it up for DI, or hooking it up to microsoft loggin.  that work is done in CreateHostBuilder
            var seriLogger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .CreateLogger();
            try
            {
                seriLogger.Information("Starting app...");
                var builder = Host.CreateDefaultBuilder();

                //tell the builder that we want to use our configuration that we initialized from appsettings
                builder.ConfigureAppConfiguration(x => x.AddConfiguration(configuration));

                //we put configure services into a separate method because its so verbose
                InitServices(builder);

                //create the host, after this point, builder cannot be used because it will have been built
                host = builder.Build();

                //Save our service provider so we can use it later 
                ServiceProvider = host.Services;
            }
            catch (Exception ex)
            {
                //use seriLogger because if host builder failed, then microsoft logger is not ready
                seriLogger.Error(ex, "bad news bro");
                host = null;
            }
            return host;
        }

        static IConfigurationRoot CreateConfigurationRoot()
        {
            //tell our app it's configuation comes from appsettings.json
            return new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile("appsettings.json")
                  .Build();
        }

        static IHostBuilder InitServices(IHostBuilder builder)
        {
            builder.ConfigureServices((ctx, services) =>
              {
                  services.AddTransient<Services.ExcelParse.Abstractions.IExcelParser, Services.ExcelParse.ExcelParser>();
                  services.AddTransient<Services.SqlServerWrite.Abstractions.ISqlServerWriter, Services.SqlServerWrite.SqlServerWriter>();

                  //external packages
                  services.AddTransient<FaultlessExecution.Abstractions.IFaultlessExecutionService, FaultlessExecution.FaultlessExecutionService>();
              })
             //useserilog plugs serilog in so that we can user Microsoft.Extensions.Logger for logging
             .UseSerilog((context, services, loggerConfiguration) => loggerConfiguration
                    .ReadFrom.Configuration(context.Configuration)//this is what lets serilog get it's settings from appsetting
                    .Enrich.FromLogContext());


            return builder;
        }


        static void Run(string fileName)
        {
            //if we get here, we did not run into any issues getting everything setup. yay. party.

            var logger = Program.ServiceProvider.GetRequiredService<ILogger<Program>>();
            var parser = Program.ServiceProvider.GetRequiredService<Services.ExcelParse.Abstractions.IExcelParser>();
            var sqlWriter = Program.ServiceProvider.GetRequiredService<Services.SqlServerWrite.Abstractions.ISqlServerWriter>();
            var faultlessExecutionService = Program.ServiceProvider.GetRequiredService<FaultlessExecution.Abstractions.IFaultlessExecutionService>();

            //authenticate and store token
            logger.LogInformation("begin parsing...");

            using (var file = System.IO.File.OpenRead(fileName))
            {
                var parserSettings = Configuration.GetSection("ParserSettings").Get<Services.ExcelParse.ParseSettings>();
                var sqlSettings = Configuration.GetSection("sqlSettings").Get<Services.SqlServerWrite.SqlSettings>();
                var parseResult = faultlessExecutionService.TryExecute(() => parser.ParseWorkbook(file, parserSettings));

                if (parseResult.WasSuccessful)
                {
                    logger.LogInformation("Parse complete.");
                    if (parseResult.ReturnValue.Records.Any())
                    {
                        var sqlResult = faultlessExecutionService.TryExecute(() => sqlWriter.WriteToSqlServer(parseResult.ReturnValue.Records, sqlSettings));
                        if (sqlResult.WasSuccessful)
                        {
                            logger.LogInformation("sql writing complete.");
                        }
                        else
                        {
                            logger.LogInformation("sql writing failed. you are a failure.");
                        }
                    }
                    else
                    {
                        logger.LogInformation("There were no records to write to sql");
                    }


                    int recCount = parseResult.ReturnValue.Records.Count;
                    int warnCount = parseResult.ReturnValue.Warnings.Count;
                    int errorCount = parseResult.ReturnValue.Errors.Count;

                    parseResult.ReturnValue.Warnings.ForEach(x => logger.LogWarning(x));
                    parseResult.ReturnValue.Errors.ForEach(x => logger.LogError(x));

                    logger.LogInformation($"{warnCount} Warnings");
                    logger.LogInformation($"{errorCount} Errors");
                    logger.LogInformation($"{recCount} Total Records");
                }
                else
                {
                    logger.LogInformation("Parsing Errord. you are a failure.");
                }
            }
            logger.LogInformation("Done");
        }


    }
}
