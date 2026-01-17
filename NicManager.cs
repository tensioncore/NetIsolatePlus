﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using NetIsolatePlus.Models;

namespace NetIsolatePlus
{
    public class NicManager
    {
        public NicInfo? CurrentIsolated { get; set; }

        public List<NicInfo> ListAdapters()
        {
            var list = new List<NicInfo>();

            using var searcher = new ManagementObjectSearcher(
                "SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID IS NOT NULL");

            foreach (ManagementObject mo in searcher.Get())
            {
                var name = (mo["NetConnectionID"] as string) ?? (mo["Name"] as string) ?? "(unnamed)";
                var desc = (mo["Name"] as string) ?? "";

                var guidRaw = (mo["GUID"] as string) ?? "";
                var guid = TryNormalizeGuid(guidRaw, out var g) ? g : "";

                // GUID-ONLY ENFORCEMENT:
                // If the adapter doesn't have a valid GUID (NetCfgInstanceId), do not surface it in the app.
                // This guarantees stable identity and prevents any accidental fallback to renameable fields.
                if (string.IsNullOrWhiteSpace(guid))
                    continue;

                var deviceId = Convert.ToString(mo["DeviceID"]) ?? "";
                var pnp = (mo["PNPDeviceID"] as string) ?? "";

                // GUID is the ONLY identity used by logic
                var id = guid;

                bool? netEnabledRaw = null;
                try { if (mo["NetEnabled"] is bool b) netEnabledRaw = b; } catch { }

                uint? cmError = null;
                try
                {
                    var v = mo["ConfigManagerErrorCode"];
                    if (v != null)
                    {
                        if (v is uint u) cmError = u;
                        else if (uint.TryParse(v.ToString(), out var parsed)) cmError = parsed;
                    }
                }
                catch { }

                int? netConnStatus = null;
                try
                {
                    var v = mo["NetConnectionStatus"];
                    if (v != null)
                    {
                        if (v is int i) netConnStatus = i;
                        else if (int.TryParse(v.ToString(), out var parsed)) netConnStatus = parsed;
                    }
                }
                catch { }

                bool? physical = null;
                try { if (mo["PhysicalAdapter"] is bool pb) physical = pb; } catch { }

                // AREA 51: do not touch this logic
                var enabled = DetermineAdminEnabled(netEnabledRaw, cmError, netConnStatus);

                list.Add(new NicInfo
                {
                    Id = id,
                    Guid = guid,
                    DeviceId = deviceId,
                    PnpDeviceId = pnp,

                    Name = name,
                    Description = desc,

                    Enabled = enabled,
                    NetEnabledRaw = netEnabledRaw,
                    NetConnectionStatus = netConnStatus,
                    ConfigManagerErrorCode = cmError,
                    PhysicalAdapter = physical
                });
            }

            if (CurrentIsolated != null)
            {
                var match = list.FirstOrDefault(n => string.Equals(n.Id, CurrentIsolated.Id, StringComparison.OrdinalIgnoreCase));
                if (match != null) CurrentIsolated = match;
                else CurrentIsolated = null; // adapter vanished or no longer valid under GUID-only rules
            }

            return list;
        }

        private static bool DetermineAdminEnabled(bool? netEnabledRaw, uint? configManagerErrorCode, int? netConnectionStatus)
        {
            if (configManagerErrorCode.HasValue && configManagerErrorCode.Value == 22)
                return false;

            if (netEnabledRaw.HasValue && netEnabledRaw.Value)
                return true;

            if (netConnectionStatus.HasValue)
                return true;

            return true;
        }

        public Dictionary<string, bool> CaptureStates()
            => ListAdapters().ToDictionary(n => n.Id, n => n.Enabled, StringComparer.OrdinalIgnoreCase);

        public void RestoreStates(Dictionary<string, bool> states)
        {
            var now = ListAdapters();
            foreach (var nic in now)
            {
                if (!states.TryGetValue(nic.Id, out var shouldEnable)) continue;

                if (shouldEnable && !nic.Enabled) Enable(nic.Id);
                else if (!shouldEnable && nic.Enabled) Disable(nic.Id);
            }
        }

        public void IsolateTo(string nicId)
        {
            var adapters = ListAdapters();
            foreach (var nic in adapters)
            {
                if (string.Equals(nic.Id, nicId, StringComparison.OrdinalIgnoreCase))
                {
                    if (!nic.Enabled) Enable(nic.Id);
                    CurrentIsolated = nic;
                }
                else
                {
                    if (nic.Enabled) Disable(nic.Id);
                }
            }
        }

        public void EnablePublic(string nicId) => Enable(nicId);
        public void DisablePublic(string nicId) => Disable(nicId);

        private static void Enable(string nicId)
        {
            var mo = FindAdapter(nicId) ?? throw new InvalidOperationException("Adapter not found.");
            var result = mo.InvokeMethod("Enable", null);
            ValidateResult(result, true);
        }

        private static void Disable(string nicId)
        {
            var mo = FindAdapter(nicId) ?? throw new InvalidOperationException("Adapter not found.");
            var result = mo.InvokeMethod("Disable", null);
            ValidateResult(result, false);
        }

        private static ManagementObject? FindAdapter(string nicId)
        {
            // GUID-ONLY ENFORCEMENT:
            // We only match adapters by GUID (NetCfgInstanceId).
            // No fallback to DeviceID/PNP/name, ever.
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter");
            foreach (ManagementObject mo in searcher.Get())
            {
                var guidRaw = mo["GUID"] as string;
                if (!TryNormalizeGuid(guidRaw, out var guid)) continue;

                if (!string.IsNullOrWhiteSpace(guid) &&
                    string.Equals(nicId, guid, StringComparison.OrdinalIgnoreCase))
                    return mo;
            }
            return null;
        }

        private static void ValidateResult(object? result, bool enableOp)
        {
            uint code = 0;

            if (result is uint u) code = u;
            else if (result is ManagementBaseObject mbo)
            {
                var rv = mbo.Properties["ReturnValue"]?.Value;
                if (rv is uint ru) code = ru;
                else if (rv != null && uint.TryParse(rv.ToString(), out var parsed)) code = parsed;
            }
            else if (result == null) code = 0;

            if (code != 0)
                throw new InvalidOperationException($"Failed to {(enableOp ? "enable" : "disable")} adapter. WMI code: {code}");
        }

        public void OpenStatus(NicInfo nic)
        {
            const string ConnectionsFolder = "shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}";

            object? shellObj = null;
            object? folderObj = null;
            object? matchObj = null;
            object? itemsObj = null;

            try
            {
                var shellType = Type.GetTypeFromProgID("Shell.Application")
                                ?? throw new InvalidOperationException("Shell not available.");

                shellObj = Activator.CreateInstance(shellType)!;
                dynamic shell = shellObj;
                folderObj = shell.NameSpace(ConnectionsFolder);
                dynamic folder = folderObj;

                if (folderObj == null)
                    throw new InvalidOperationException("Cannot open Network Connections.");

                dynamic? matchItem = null;

                // Enumerate items once; explicitly release non-matching COM objects.
                itemsObj = folder.Items();
                dynamic items = itemsObj;

                if (!string.IsNullOrWhiteSpace(nic.Guid))
                {
                    foreach (var item in items)
                    {
                        object? itemObj = item; // nullable: we intentionally null it out when keeping the match
                        try
                        {
                            var itemGuid = TryGetShellItemGuid(item);
                            if (!string.IsNullOrWhiteSpace(itemGuid) &&
                                string.Equals(itemGuid, nic.Guid, StringComparison.OrdinalIgnoreCase))
                            {
                                matchItem = item;
                                matchObj = itemObj;
                                itemObj = null; // don't release match
                                break;
                            }
                        }
                        finally
                        {
                            if (itemObj != null) ReleaseCom(itemObj);
                        }
                    }
                }

                if (matchItem == null)
                {
                    foreach (var item in items)
                    {
                        object? itemObj = item; // nullable: we intentionally null it out when keeping the match
                        try
                        {
                            string itemName = Convert.ToString(item.Name) ?? "";
                            if (string.Equals(itemName, nic.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                matchItem = item;
                                matchObj = itemObj;
                                itemObj = null; // don't release match
                                break;
                            }
                        }
                        finally
                        {
                            if (itemObj != null) ReleaseCom(itemObj);
                        }
                    }
                }

                if (matchItem == null)
                {
                    Process.Start(new ProcessStartInfo("control.exe", "ncpa.cpl") { UseShellExecute = true });
                    return;
                }

                if (TryInvokeVerbAndConfirmDialog(matchItem, "status", nic.Name)) return;

                try
                {
                    // Enumerate verbs and release each COM verb object.
                    foreach (var verb in matchItem.Verbs())
                    {
                        object? verbObj = verb;
                        try
                        {
                            string vname = (Convert.ToString(verb.Name) ?? "").Replace("&", "").Trim();
                            if (!vname.StartsWith("Status", StringComparison.OrdinalIgnoreCase)) continue;

                            try
                            {
                                verb.DoIt();
                                if (WaitForStatusOrProperties(nic.Name, 450)) return;
                            }
                            catch { }
                        }
                        finally
                        {
                            if (verbObj != null) ReleaseCom(verbObj);
                        }
                    }
                }
                catch { }

                if (TryInvokeVerbAndConfirmDialog(matchItem, "open", nic.Name)) return;
                if (TryInvokeVerbAndConfirmDialog(matchItem, "properties", nic.Name)) return;

                try
                {
                    matchItem.InvokeVerb();
                    if (WaitForStatusOrProperties(nic.Name, 450)) return;
                }
                catch { }

                Process.Start(new ProcessStartInfo("control.exe", "ncpa.cpl") { UseShellExecute = true });
            }
            catch
            {
                try { Process.Start(new ProcessStartInfo("ms-settings:network-status") { UseShellExecute = true }); }
                catch { Process.Start(new ProcessStartInfo("control.exe", "ncpa.cpl") { UseShellExecute = true }); }
            }
            finally
            {
                // Release in reverse-ish order; never throw from cleanup.
                if (matchObj != null) ReleaseCom(matchObj);
                if (itemsObj != null) ReleaseCom(itemsObj);
                if (folderObj != null) ReleaseCom(folderObj);
                if (shellObj != null) ReleaseCom(shellObj);
            }
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }

        private static bool TryInvokeVerbAndConfirmDialog(dynamic item, string verb, string nicName)
        {
            try
            {
                item.InvokeVerb(verb);
                return WaitForStatusOrProperties(nicName, 450);
            }
            catch
            {
                return false;
            }
        }

        private static bool WaitForStatusOrProperties(string nicName, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (ExternalWindowPlacer.FindStatusOrPropertiesHandle(nicName) != IntPtr.Zero)
                    return true;

                Thread.Sleep(50);
            }
            return false;
        }

        private static string? TryGetShellItemGuid(dynamic item)
        {
            try
            {
                string? v = Convert.ToString(item.ExtendedProperty("System.NetworkInterface.GUID"));
                if (TryNormalizeGuid(v, out var g)) return g;
            }
            catch { }

            try
            {
                string? v = Convert.ToString(item.ExtendedProperty("System.DeviceInterface.GUID"));
                if (TryNormalizeGuid(v, out var g)) return g;
            }
            catch { }

            foreach (var candidate in new[]
            {
                SafeToString(() => item.Path),
                SafeToString(() => item.Self?.Path),
                SafeToString(() => item.ExtendedProperty("System.ItemUrl")),
                SafeToString(() => item.ExtendedProperty("System.ItemUrlType")),
            })
            {
                var g = ExtractGuid(candidate);
                if (g != null) return g;
            }

            return null;
        }

        private static string SafeToString(Func<object?> getter)
        {
            try { return Convert.ToString(getter()) ?? ""; } catch { return ""; }
        }

        private static string? ExtractGuid(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            var m = Regex.Match(s, @"\{[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}\}");
            if (!m.Success) return null;
            return TryNormalizeGuid(m.Value, out var g) ? g : null;
        }

        // Unified GUID validator + normalizer:
        // Accepts "GUID" or "{GUID}" and outputs normalized "D" format.
        private static bool TryNormalizeGuid(string? s, out string normalized)
        {
            normalized = "";
            if (string.IsNullOrWhiteSpace(s)) return false;

            s = s.Trim();
            if (!s.StartsWith("{")) s = "{" + s + "}";

            if (!Guid.TryParse(s, out var g)) return false;

            normalized = g.ToString("D");
            return true;
        }
    }
}
