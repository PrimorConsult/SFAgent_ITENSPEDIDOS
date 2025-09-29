using System;
using System.Diagnostics;
using System.IO;

namespace SFAgent.Utils
{
    public static class Logger
    {
        private static readonly object _sync = new object();
        private static readonly string LogDir = @"C:\SFAgent\Logs - ItensPedidos";
        private static readonly string LogFile = Path.Combine(LogDir, "LogService_ITENSPEDIDOS.log");

        /// <summary>
        /// Inicializa o log, limpando qualquer conteúdo anterior.
        /// Deve ser chamado no OnStart do serviço.
        /// </summary>
        public static void InitLog()
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                File.WriteAllText(LogFile,
                    $"=== Nova execução iniciada em {DateTime.Now:dd/MM/yyyy HH:mm:ss} ==={Environment.NewLine}");
            }
            catch
            {
                // nunca quebrar por erro de log
            }
        }

        /// <summary>
        /// Registra uma linha de log no arquivo e no Event Viewer (apenas erros).
        /// </summary>
        public static void Log(string message, bool asError = false)
        {
            var line = $"{DateTime.Now:dd/MM/yyyy HH:mm:ss} - {Environment.MachineName} - {message}";

            try
            {
                lock (_sync)
                {
                    Directory.CreateDirectory(LogDir);

                    // adiciona a linha ao arquivo (não substitui)
                    File.AppendAllText(LogFile, line + Environment.NewLine);
                }

                // somente erros vão para o Event Viewer
                if (asError)
                {
                    EnsureEventSource();
                    EventLog.WriteEntry("SFAgent", message, EventLogEntryType.Error);
                }
            }
            catch
            {
                // nunca quebrar por erro de log
            }
        }

        private static void EnsureEventSource()
        {
            if (!EventLog.SourceExists("SFAgent"))
                EventLog.CreateEventSource("SFAgent", "Application");
        }
    }
}
