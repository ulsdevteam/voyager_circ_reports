using CommandLine;

namespace voyager_circ_reports
{
    class Program
    {        
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(options => {
                    var report = Report.Run(options);
                    if (options.Output is null) {
                        System.Console.WriteLine(report);
                    } else {
                        report.ExportToExcel(options.Output);
                    }
                });
        }
    }
}
