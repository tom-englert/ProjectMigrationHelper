namespace ProjectMigrationHelper
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;

    using JetBrains.Annotations;

    using Newtonsoft.Json.Linq;

    internal class DteSolution
    {
        private readonly ITracer _tracer;

        public DteSolution(EnvDTE80.DTE2 dte, ITracer tracer)
        {
            Dte = dte;
            _tracer = tracer;
        }

        public IDictionary<string, JObject> CreateFingerprints()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var items = new Dictionary<string, JObject>();

            try
            {
                var index = 0;

                foreach (var project in GetProjects().OrderBy(item => item.Name))
                {
                    var name = @"<unknown>";

                    try
                    {
                        index += 1;
                        name = project.Name;

                        _tracer.WriteLine("Reading project: " + name);

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

        [CanBeNull]
        public string Folder => Path.GetDirectoryName(Solution?.FullName);

        // ReSharper disable once SuspiciousTypeConversion.Global
        [CanBeNull]
        private EnvDTE80.Solution2 Solution => (EnvDTE80.Solution2)Dte.Solution;

        private EnvDTE80.DTE2 Dte { get; }

        private IEnumerable<EnvDTE.Project> GetProjects()
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
    }
}
