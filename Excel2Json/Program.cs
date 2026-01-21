using System;
using System.CommandLine;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2Json
{
    internal class Program
    {
        static readonly Option<string> inputOption = new(
             name: "--input",
             aliases: "-i") {
            Required = true,
            Description = "directory or file",
        };

        static readonly Option<string> outputOption = new(
            name: "--output",
            aliases: "-o") {
            Required = true,
            Description = "directory",
        };

        static async Task Run(ParseResult parseResult)
        {
            string srcPath = parseResult.GetRequiredValue(inputOption);
            string dstDir = parseResult.GetRequiredValue(outputOption);

            string[] srcFiles = null;
            string[] dstFiles = null;

            bool isValid = false;

            if (!Directory.Exists(dstDir)) {
                throw new ArgumentException("Dst Directory Not Exists !");
            }

            if (File.Exists(srcPath)) {
                srcFiles = new string[] { srcPath };
                dstFiles = new string[] { Path.Combine(dstDir, Path.GetFileName(Path.ChangeExtension(srcPath, "json"))) };
                isValid = true;
            }
            else if (Directory.Exists(srcPath)) {
                const string pattern = @"^[a-zA-Z].*$";
                srcFiles = Directory.GetFiles(srcPath, "*.xls*").Where(fileName => Regex.IsMatch(Path.GetFileName(fileName), pattern)).ToArray();
                dstFiles = new string[srcFiles.Length];
                for (int i = 0; i < srcFiles.Length; ++i) {
                    dstFiles[i] = Path.Combine(dstDir, Path.GetFileName(Path.ChangeExtension(srcFiles[i], "json")));
                }
                isValid = true;
            }

            if (isValid) {
                Parallel.For(0, srcFiles.Length, i => {
                    ExcelHandler.Handle(srcFiles[i], dstFiles[i]);
                });
            }
            else {
                throw new ArgumentException("Src File or Directory Not Exists !");
            }
            await Task.Yield();
        }

        static void Main(string[] args)
        {
            var rootCommand = new RootCommand("Excel To Json Tool");
            rootCommand.Add(inputOption);
            rootCommand.Add(outputOption);
            rootCommand.SetAction(Run);
            rootCommand.Parse(args).Invoke();
        }
    }
}
