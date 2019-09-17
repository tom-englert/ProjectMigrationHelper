namespace ProjectMigrationHelper
{
    using System.ComponentModel;
    using System.Globalization;

    using JetBrains.Annotations;

    public interface ITracer
    {
        void TraceError([Localizable(false)][NotNull] string value);

        void TraceWarning([Localizable(false)][NotNull] string value);

        void WriteLine([Localizable(false)][NotNull] string value);
    }

    public static class TracerExtensions
    {
        [StringFormatMethod("format")]
        public static void TraceError([NotNull] this ITracer tracer, [Localizable(false)][NotNull] string format, [NotNull][ItemNotNull] params object[] args)
        {
            tracer.TraceError(string.Format(CultureInfo.CurrentCulture, format, args));
        }

        [StringFormatMethod("format")]
        public static void TraceWarning([NotNull] this ITracer tracer, [Localizable(false)][NotNull] string format, [NotNull][ItemNotNull] params object[] args)
        {
            tracer.TraceWarning(string.Format(CultureInfo.CurrentCulture, format, args));
        }

        [StringFormatMethod("format")]
        public static void WriteLine([NotNull] this ITracer tracer, [Localizable(false)][NotNull] string format, [NotNull][ItemNotNull] params object[] args)
        {
            tracer.WriteLine(string.Format(CultureInfo.CurrentCulture, format, args));
        }
    }
}