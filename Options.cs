using CommandLine;

namespace voyager_circ_reports 
{
    public class Options
    {
        [Value(0, Required = true, MetaName = "Month", HelpText = "Month between 1 (Jan) and 12 (Dec)")]
        public int Month { get; set; }
        [Value(1, Required = true, MetaName = "Year")]
        public int Year { get; set; }
        [Option('l', "location", HelpText = "Library location code")]
        public string Location { get; set; }
        [Option('o', "output", HelpText = "File path to write Excel spreadsheet")]      
        public string Output { get; set; }
        [Option('c', "conn", Required = true, HelpText = "Oracle connection string")]
        public string ConnectionString { get; set; }
    }
}
