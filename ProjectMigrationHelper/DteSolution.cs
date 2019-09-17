namespace ProjectMigrationHelper
{
    using System;
    using System.Collections.Generic;
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

        /// <summary>
        /// Gets all files of all project in the solution.
        /// </summary>
        /// <returns>The files.</returns>
        [NotNull, ItemNotNull]
        public JObject CreateFingerprint()
        {
            var jSolution = new JObject();

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

                        var jProject = new JObject();

                        jSolution.Add(name, jProject);

                        AddProjectItems(name, project.ProjectItems, jProject);
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

            return jSolution;
        }

        public string Folder => Path.GetDirectoryName(Solution.FullName);

        // ReSharper disable once SuspiciousTypeConversion.Global
        private EnvDTE80.Solution2 Solution => (EnvDTE80.Solution2)Dte.Solution ?? throw new InvalidOperationException();

        private EnvDTE80.DTE2 Dte { get; }

        private IEnumerable<EnvDTE.Project> GetProjects()
        {
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
                    // _tracer.TraceError("Error loading project #" + i);
                    continue;
                }

                yield return project;
            }
        }

        private void AddProjectItems([CanBeNull] string projectName, [ItemNotNull][CanBeNull] EnvDTE.ProjectItems projectItems, [NotNull] JObject jProject)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItems == null)
                return;

            try
            {
                var index = 1;

                foreach (var projectItem in projectItems.OfType<EnvDTE.ProjectItem>().OrderBy(item => item.Name))
                {
                    try
                    {
                        AddProjectItems(projectName, projectItem, jProject);
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

        private void AddProjectItems([CanBeNull] string projectName, [NotNull] EnvDTE.ProjectItem projectItem, [NotNull] JObject jParentItem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (projectItem.Object is VSLangProj.References) // MPF project (e.g. WiX) references folder, do not traverse...
                return;

            var jProjectItem = new JObject();

            foreach (var property in projectItem.Properties.OfType<EnvDTE.Property>().OrderBy(item => item.Name))
            {
                var propertyName = "<unknown>";
                try
                {
                    propertyName = property.Name;

                    var ignored = new[] { "{", "Date", "File", "Extender", "LocalPath", "URL", "Identity" };
                    if (ignored.Any(prefix => propertyName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    // exclude properties with their default values:
                    var propertyValue = property.Value;

                    if (propertyValue is string stringValue && string.IsNullOrEmpty(stringValue))
                        continue;
                    if (propertyValue is bool booleanValue && !booleanValue)
                        continue;
                    if (propertyName == "Link" && string.Equals(projectItem.Name, propertyValue as string, StringComparison.OrdinalIgnoreCase))
                        continue;

                    jProjectItem.Add(propertyName, JToken.FromObject(propertyValue));
                }
                catch (Exception ex)
                {
                    _tracer.TraceError("Error reading property {0} in project {1}: {2}", propertyName, projectName ?? "unknown", ex.Message);
                }
            }

            jParentItem.Add(projectItem.Name, jProjectItem);

            AddProjectItems(projectName, projectItem.ProjectItems, jProjectItem);

            if (projectItem.SubProject != null)
            {
                AddProjectItems(projectName, projectItem.SubProject.ProjectItems, jProjectItem);
            }
        }
    }
}
