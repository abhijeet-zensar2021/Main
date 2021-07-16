using System.Diagnostics;

namespace HSL.NFUM.PackageDeployer.DataImport.Interfaces
{
    public interface ITraceLogging
    {
        void Log(string content, TraceEventType type);
    }
}
