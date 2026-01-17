using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace NetIsolatePlus.Services
{
    public static class StartupManager
    {
        // Scheduled Task name (root folder)
        private const string TaskName = "NetIsolatePlus";

        private static string CurrentExePath()
        {
            try
            {
                var p = Environment.ProcessPath;
                if (!string.IsNullOrWhiteSpace(p))
                    return p!;
            }
            catch { }

            try
            {
                var p = Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(p))
                    return p!;
            }
            catch { }

            // Last resort: BaseDirectory is often a folder. Only accept it if we can resolve an actual EXE.
            try
            {
                var baseDir = AppContext.BaseDirectory;
                if (!string.IsNullOrWhiteSpace(baseDir))
                {
                    var trimmed = baseDir.Trim();

                    // If it looks like a file path and exists, accept it.
                    if (File.Exists(trimmed))
                        return trimmed;

                    // If it's a directory, try "<baseDir>\<processname>.exe"
                    if (Directory.Exists(trimmed))
                    {
                        var guess = Path.Combine(trimmed, Process.GetCurrentProcess().ProcessName + ".exe");
                        if (File.Exists(guess))
                            return guess;
                    }
                }
            }
            catch { }

            return "";
        }

        private static string QuotedExe()
        {
            var exe = CurrentExePath();
            if (string.IsNullOrWhiteSpace(exe))
                return "";

            return $"\"{exe}\"";
        }

        private static string NormalizePath(string? p)
        {
            if (string.IsNullOrWhiteSpace(p)) return "";
            try
            {
                p = p.Trim().Trim('"');
                return Path.GetFullPath(p);
            }
            catch
            {
                return p.Trim().Trim('"');
            }
        }

        private static (int exitCode, string output) ExecSchtasks(string arguments)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "schtasks.exe",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using var p = new Process { StartInfo = psi };

            try
            {
                p.Start();
            }
            catch (Exception ex)
            {
                return (-1, "Failed to start schtasks.exe:\n\n" + ex.Message);
            }

            // Read stdout/stderr concurrently to avoid deadlocks if buffers fill.
            var stdoutTask = p.StandardOutput.ReadToEndAsync();
            var stderrTask = p.StandardError.ReadToEndAsync();

            // Hard timeout so we never hang the app if Task Scheduler is unresponsive.
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                p.WaitForExitAsync(cts.Token).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException)
            {
                try { p.Kill(entireProcessTree: true); } catch { }
                return (-1, "schtasks.exe timed out.");
            }
            catch (Exception ex)
            {
                try { p.Kill(entireProcessTree: true); } catch { }
                return (-1, "schtasks.exe failed:\n\n" + ex.Message);
            }

            string stdout = "";
            string stderr = "";

            try { stdout = stdoutTask.GetAwaiter().GetResult(); } catch { }
            try { stderr = stderrTask.GetAwaiter().GetResult(); } catch { }

            var output = (stdout + "\n" + stderr).Trim();
            return (p.ExitCode, output);
        }

        private static bool TaskExists(out string output)
        {
            var (code, outText) = ExecSchtasks($"/Query /TN \"{TaskName}\"");
            output = outText;
            return code == 0;
        }

        private static bool TaskIsEnabled()
        {
            if (!TaskExists(out var output))
                return false;

            // Best-effort parse: if schtasks output indicates "Disabled", treat as disabled.
            // NOTE: localized OS can change this word; we're intentionally NOT fixing localization in v1.
            if (output.IndexOf("Disabled", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;

            return true;
        }

        private static bool TryGetTaskCommandPath(out string? commandPath)
        {
            commandPath = null;

            var (code, output) = ExecSchtasks($"/Query /TN \"{TaskName}\" /XML");
            if (code != 0 || string.IsNullOrWhiteSpace(output))
                return false;

            try
            {
                var doc = XDocument.Parse(output);

                // Task Scheduler XML namespace
                XNamespace ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";

                var cmd = doc.Descendants(ns + "Command").FirstOrDefault()?.Value;
                if (string.IsNullOrWhiteSpace(cmd))
                    return false;

                commandPath = cmd.Trim();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TaskPointsToCurrentExe()
        {
            if (!TryGetTaskCommandPath(out var cmd) || string.IsNullOrWhiteSpace(cmd))
                return false;

            var taskCmd = NormalizePath(cmd);
            var curExe = NormalizePath(CurrentExePath());

            if (string.IsNullOrWhiteSpace(taskCmd) || string.IsNullOrWhiteSpace(curExe))
                return false;

            return string.Equals(taskCmd, curExe, StringComparison.OrdinalIgnoreCase);
        }

        private static void CreateOrUpdateTask()
        {
            var quoted = QuotedExe();
            if (string.IsNullOrWhiteSpace(quoted))
                throw new InvalidOperationException("Unable to resolve executable path for startup task.");

            // /IT is required for GUI apps (interactive).
            // /RL HIGHEST is required because the app manifest is requireAdministrator.
            // /SC ONLOGON is per-user logon trigger when created in the current user context.
            // /F to overwrite existing task.
            var createArgs =
                $"/Create /F " +
                $"/SC ONLOGON " +
                $"/TN \"{TaskName}\" " +
                $"/TR {quoted} " +
                $"/RL HIGHEST " +
                $"/IT";

            var (code, output) = ExecSchtasks(createArgs);
            if (code != 0)
                throw new InvalidOperationException("Failed to create/update startup task:\n\n" + output);

            // Ensure enabled
            var (code2, output2) = ExecSchtasks($"/Change /TN \"{TaskName}\" /Enable");
            if (code2 != 0)
                throw new InvalidOperationException("Created startup task but could not enable it:\n\n" + output2);
        }

        private static void DeleteOrDisableTask()
        {
            // Try delete first
            var (code, output) = ExecSchtasks($"/Delete /F /TN \"{TaskName}\"");
            if (code == 0)
                return;

            // If delete fails (policy/permissions), disable it
            var (code2, output2) = ExecSchtasks($"/Change /TN \"{TaskName}\" /Disable");
            if (code2 != 0)
                throw new InvalidOperationException(
                    "Failed to remove startup task, and disabling also failed:\n\n" +
                    output + "\n\n" + output2);
        }

        public static bool IsEnabled()
        {
            return TaskIsEnabled();
        }

        public static void SetEnabled(bool enable)
        {
            if (enable)
            {
                // No-op if already correct (avoid rewriting task constantly)
                if (TaskIsEnabled() && TaskPointsToCurrentExe())
                    return;

                CreateOrUpdateTask();
            }
            else
            {
                DeleteOrDisableTask();
            }
        }

        public static void Reconcile()
        {
            // If enabled, ensure task points to the current EXE path.
            // Do NOT rewrite unless it actually needs changing.
            if (!IsEnabled())
                return;

            if (!TaskExists(out _))
                return;

            if (!TaskPointsToCurrentExe())
                CreateOrUpdateTask();
        }
    }
}
