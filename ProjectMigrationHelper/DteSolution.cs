using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

using EnvDTE80;

using JetBrains.Annotations;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using TomsToolbox.Composition;

namespace ProjectMigrationHelper
{
    internal class DteSolution
    {
        private readonly ITracer _tracer;
        private readonly IList<EnvDTE.Project> _projects;

        public DteSolution(EnvDTE80.DTE2 dte, ITracer tracer)
        {
            Dte = dte;
            _tracer = tracer;
            _projects = EnumerateProjects();
        }

        [NotNull]
        public IDictionary<string, JObject> CreateProjectFingerprints()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var items = new Dictionary<string, JObject>();

            try
            {
                var index = 0;

                foreach (var project in _projects)
                {
                    var name = @"<unknown>";

                    try
                    {
                        index += 1;
                        name = project.Name;

                        _tracer.WriteLine("Reading project file: " + name);

                        var jProject = AddProjectItems(project.ProjectItems, new JObject());

                        items.Add(name, jProject);

                    }
                    catch (Exception ex)
                    {
                        _tracer.TraceWarning("Error loading project {0}[{1}]: {2}", name, index, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _tracer.TraceError("Error loading projects: {0}", ex);
            }

            return items;
        }

        [NotNull]
        public IDictionary<string, string> CreateMefFingerprints()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var items = new Dictionary<string, string>();

            try
            {
                var index = 0;

                foreach (var project in _projects)
                {
                    var name = @"<unknown>";

                    try
                    {
                        index += 1;
                        name = project.Name;

                        if (!(project.Object is VSLangProj.VSProject))
                            continue;

                        _tracer.WriteLine("Reading MEF attributes from project: " + name);

                        var activeConfiguration = project.ConfigurationManager?.ActiveConfiguration;
                        if (activeConfiguration == null)
                            continue;

                        var primaryOutputFileName = activeConfiguration.OutputGroups?.Item("Built").GetFileNames()?.FirstOrDefault();
                        var properties = activeConfiguration.Properties;
                        var outputDirectory = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(project.FullName), properties?.Item(@"OutputPath")?.Value as string ?? string.Empty));
                        var primaryOutputFilePath = Path.Combine(outputDirectory, primaryOutputFileName);

                        if (primaryOutputFilePath == null || !File.Exists(primaryOutputFilePath))
                        {
                            _tracer.TraceWarning("Skip analysis for project, primary output does not exist: {0}[{1}]@{2}=>{3}", name, index, activeConfiguration.ToDisplayName(), primaryOutputFileName ?? "<unknown>");
                            continue;
                        }

                        var assembly = Assembly.LoadFrom(primaryOutputFilePath);

                        var metadata = MetadataReader.Read(assembly);

                        if (!metadata.Any())
                        {
                            _tracer.WriteLine("Project skipped, no MEF annotations found: {0}[{1}]@{2}=>{3}", name, index, activeConfiguration.ToDisplayName(), primaryOutputFileName ?? "<unknown>");
                            continue;
                        }

                        var data = Serialize(metadata);

                        items.Add(name, data);

                    }
                    catch (Exception ex)
                    {
                        _tracer.TraceWarning("Error loading project {0}[{1}]: {2}", name, index, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _tracer.TraceError("Error loading projects: {0}", ex);
            }

            return items;
        }

        private static string Serialize(IList<ExportInfo> result)
        {
            return JsonConvert.SerializeObject(result, Formatting.Indented);
        }


        [CanBeNull]
        public string Folder => Path.GetDirectoryName(Solution?.FullName);

        // ReSharper disable once SuspiciousTypeConversion.Global
        [CanBeNull]
        private EnvDTE80.Solution2 Solution => (EnvDTE80.Solution2)Dte.Solution;

        private EnvDTE80.DTE2 Dte { get; }

        [NotNull]
        private JObject AddProjectItems([ItemNotNull][CanBeNull] EnvDTE.ProjectItems projectItems, [NotNull] JObject jProject)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItems == null)
                return jProject;

            try
            {
                var index = 1;

                foreach (var projectItem in projectItems.OfType<EnvDTE.ProjectItem>().OrderBy(item => item.Name))
                {
                    try
                    {
                        AddProjectItems(projectItem, jProject);
                    }
                    catch
                    {
                        _tracer.TraceWarning("Can't load project item #{0}", index);
                    }

                    index += 1;
                }
            }
            catch
            {
                _tracer.TraceWarning("Can't load a project item.");
            }

            return jProject;
        }

        private void AddProjectItems([NotNull] EnvDTE.ProjectItem projectItem, [NotNull] JObject jParentItem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItem.Object is VSLangProj.References) // MPF project (e.g. WiX) references folder, do not traverse...
                return;

            if (projectItem.Properties == null)
                return;

            var jProjectItem = new JObject();

            foreach (var property in projectItem.Properties.OfType<EnvDTE.Property>().OrderBy(item => item.Name))
            {
                var propertyName = "<unknown>";
                try
                {
                    propertyName = property.Name;

                    var ignored = new[] { "{", "Date", "File", "Extender", "LocalPath", "URL", "Identity", "BuildAction", "SubType" };
                    if (ignored.Any(prefix => propertyName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    // exclude properties with their default values:
                    var propertyValue = property.Value;
                    if (propertyValue == null)
                        continue;
                    if (propertyValue is string stringValue && string.IsNullOrEmpty(stringValue))
                        continue;
                    if (propertyValue is bool booleanValue && !booleanValue)
                        continue;
                    if (((propertyName == "Link") || (propertyName == "FolderName"))
                        && string.Equals(projectItem.Name, propertyValue as string, StringComparison.OrdinalIgnoreCase))
                        continue;

                    jProjectItem.Add(propertyName, JToken.FromObject(propertyValue));
                }
                catch (Exception ex)
                {
                    _tracer.TraceWarning("Can't read property {0}: {1}", propertyName, ex.Message);
                }
            }

            jParentItem.Add(projectItem.Name, jProjectItem);

            AddProjectItems(projectItem.ProjectItems, jProjectItem);

            if (projectItem.SubProject != null)
            {
                AddProjectItems(projectItem.SubProject.ProjectItems, jProjectItem);
            }
        }

        private IList<EnvDTE.Project> EnumerateProjects()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var items = new List<EnvDTE.Project>();

            try
            {
                var index = 0;

                foreach (var project in GetRootProjects())
                {
                    var name = @"<unknown>";

                    try
                    {
                        index += 1;
                        name = project.Name;

                        _tracer.WriteLine("Load project: {0}", name);

                        items.Add(project);

                        GetSubProjects(name, project.ProjectItems, items);
                    }
                    catch (Exception ex)
                    {
                        _tracer.TraceWarning("Error loading project {0}[{1}]: {2}", name, index, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _tracer.TraceError("Error loading projects: {0}", ex);
            }

            return items.Where(item => item.Kind != ProjectKinds.vsProjectKindSolutionFolder).ToList().AsReadOnly();
        }

        private IEnumerable<EnvDTE.Project> GetRootProjects()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var solution = Solution;

            var projects = solution?.Projects;
            if (projects == null)
                yield break;

            for (var i = 1; i <= projects.Count; i++)
            {
                EnvDTE.Project project;
                try
                {
                    project = projects.Item(i);
                }
                catch
                {
                    _tracer.TraceError("Error loading project #" + i);
                    continue;
                }

                yield return project;
            }
        }

        private void GetSubProjects(string projectName, EnvDTE.ProjectItems projectItems, IList<EnvDTE.Project> items)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItems == null)
                return;

            try
            {
                var index = 1;

                foreach (var projectItem in projectItems.OfType<EnvDTE.ProjectItem>())
                {
                    try
                    {
                        GetSubProjects(projectName, projectItem, items);
                    }
                    catch
                    {
                        _tracer.TraceError("Error loading project item #{0} in project {1}.", index, projectName ?? "unknown");
                    }

                    index += 1;
                }
            }
            catch
            {
                _tracer.TraceError("Error loading a project item in project {0}.", projectName ?? "unknown");
            }
        }

        private void GetSubProjects(string projectName, EnvDTE.ProjectItem projectItem, IList<EnvDTE.Project> items)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItem.Object is VSLangProj.References) // MPF project (e.g. WiX) references folder, do not traverse...
                return;

            if (projectItem.Object is EnvDTE.Project project)
            {
                var name = project.Name;
                _tracer.WriteLine("Load project: {0}", name);
                items.Add(project);
            }

            GetSubProjects(projectName, projectItem.ProjectItems, items);

            if (projectItem.SubProject != null)
            {
                GetSubProjects(projectName, projectItem.SubProject.ProjectItems, items);
            }
        }
    }

    internal static class ExtensionMethods
    {
        [CanBeNull]
        public static IEnumerable<string> GetFileNames([CanBeNull] this EnvDTE.OutputGroup outputGroup)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                return ((Array)outputGroup?.FileNames)?.OfType<string>();
            }
            catch
            {
                return null;
            }
        }

        [NotNull]
        public static string ToDisplayName([CanBeNull] this EnvDTE.Configuration configuration)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (configuration == null)
                return "<undefined>";

            return configuration.ConfigurationName + "|" + configuration.PlatformName;

        }
    }
}
