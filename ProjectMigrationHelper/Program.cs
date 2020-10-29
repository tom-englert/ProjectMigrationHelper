using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

using Newtonsoft.Json;

using TomsToolbox.Composition;

namespace ProjectMigrationHelper
{
    static class Program
    {
        public static void Main()
        {
            try
            {
                using (var inputReader = new StreamReader(Console.OpenStandardInput()))
                {
                    var input = inputReader.ReadToEnd();

                    var assemblies = JsonConvert.DeserializeObject<Dictionary<string, string>>(input);

                    var metadata = ReadMetadata(assemblies);

                    var output = JsonConvert.SerializeObject(metadata);

                    Console.Write(output);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private static IDictionary<string, string> ReadMetadata(IDictionary<string, string> assemblies)
        {
            var items = new Dictionary<string, string>();

            foreach (var assembly in assemblies)
            {
                var fingerPrint = ReadMetadata(assembly.Value);
                if (!string.IsNullOrEmpty(fingerPrint))
                {
                    items.Add(assembly.Key, fingerPrint);
                }
            }

            return items;
        }

        private static string ReadMetadata(string primaryOutputFilePath)
        {
            var assembly = Assembly.LoadFrom(primaryOutputFilePath);
            var metadata = MetadataReader.Read(assembly);
            if (!metadata.Any())
                return null;

            return Serialize(metadata.OrderBy(m => m.Type?.FullName).ToList());
        }

        private static string Serialize(IList<ExportInfo> result)
        {
            return JsonConvert.SerializeObject(result, Formatting.Indented);
        }
    }
}
