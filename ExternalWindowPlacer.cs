﻿using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace NetIsolatePlus
{
    internal static class ExternalWindowPlacer
    {
        // Existing: used by NicManager confirmation checks
        internal static IntPtr FindStatusWindowHandle(string nicName) => FindStatusWindow(nicName);

        // Existing: used by NicManager confirmation checks
        internal static IntPtr FindPropertiesWindowHandle(string nicName) => FindPropertiesWindow(nicName);

        // Convenience: status first, else properties
        internal static IntPtr FindStatusOrPropertiesHandle(string nicName)
        {
            var h = FindStatusWindow(nicName);
            if (h != IntPtr.Zero) return h;
            return FindPropertiesWindow(nicName);
        }

        // Center whichever dialog is present (Status preferred; else Properties)
        public static async Task CenterStatusWindowAsync(Window owner, string nicName, int timeoutMs = 3000)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var h = FindStatusOrPropertiesHandle(nicName);
                if (h != IntPtr.Zero)
                {
                    GetWindowRect(h, out var r);
                    int w = r.Right - r.Left, hgt = r.Bottom - r.Top;
                    int x = (int)owner.Left + (int)((owner.Width - w) / 2);
                    int y = (int)owner.Top + (int)((owner.Height - hgt) / 2);
                    SetWindowPos(h, IntPtr.Zero, x, y, 0, 0, 0x0001 | 0x0004 | 0x0010); // NOSIZE|NOZORDER|NOACTIVATE
                    return;
                }
                await Task.Delay(100);
            }
        }

        // --- NEW CORE MATCHER ---
        // Don't rely on English words ("Status"/"Properties").
        // Most builds title dialogs like "<NIC NAME> <localized word>" so NIC name is stable even on non-English Windows.
        // Also require dialog class "#32770" to avoid matching random windows.
        private static IntPtr FindDialogByNicName(string nicName)
        {
            if (string.IsNullOrWhiteSpace(nicName)) return IntPtr.Zero;

            // Pass 1: prefer titles that START with the NIC name (reduces accidental matches)
            var found = FindDialogByNicNameCore(nicName, requireStartsWith: true);
            if (found != IntPtr.Zero) return found;

            // Pass 2: fallback to "contains" match (best effort)
            return FindDialogByNicNameCore(nicName, requireStartsWith: false);
        }

        private static IntPtr FindDialogByNicNameCore(string nicName, bool requireStartsWith)
        {
            IntPtr found = IntPtr.Zero;

            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;

                // Must be a dialog window
                var cls = GetClass(h);
                if (!string.Equals(cls, "#32770", StringComparison.Ordinal)) return true;

                var title = GetText(h);
                if (string.IsNullOrWhiteSpace(title)) return true;

                var t = title.Trim();

                if (requireStartsWith)
                {
                    // Title usually starts with NIC name (e.g. "<NIC> Status", "<NIC> Properties", localized suffix)
                    if (t.StartsWith(nicName, StringComparison.OrdinalIgnoreCase))
                    {
                        found = h;
                        return false;
                    }
                    return true;
                }

                // Fallback: any contains match
                if (t.IndexOf(nicName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = h;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        // Legacy status finder now prefers the dialog matcher, then falls back to old keyword logic.
        private static IntPtr FindStatusWindow(string nicName)
        {
            // Prefer localization-safe detection
            var dlg = FindDialogByNicName(nicName);
            if (dlg != IntPtr.Zero) return dlg;

            // Fallback: English keyword (best effort)
            IntPtr found = IntPtr.Zero;

            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;

                var t = GetText(h);
                if (string.IsNullOrWhiteSpace(t)) return true;

                var tn = t.Trim();
                if (tn.IndexOf("Status", StringComparison.OrdinalIgnoreCase) < 0) return true;

                bool matchesNic =
                    (!string.IsNullOrWhiteSpace(nicName) &&
                     tn.IndexOf(nicName, StringComparison.OrdinalIgnoreCase) >= 0);

                if (matchesNic || tn.EndsWith(" Status", StringComparison.OrdinalIgnoreCase))
                {
                    found = h;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        // Legacy properties finder now prefers the dialog matcher, then falls back to old keyword logic.
        private static IntPtr FindPropertiesWindow(string nicName)
        {
            // Prefer localization-safe detection
            var dlg = FindDialogByNicName(nicName);
            if (dlg != IntPtr.Zero) return dlg;

            // Fallback: English keyword (best effort)
            IntPtr found = IntPtr.Zero;

            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;

                var t = GetText(h);
                if (string.IsNullOrWhiteSpace(t)) return true;

                var tn = t.Trim();

                bool looksLikeProps =
                    tn.IndexOf("Properties", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!looksLikeProps) return true;

                bool matchesNic =
                    (!string.IsNullOrWhiteSpace(nicName) &&
                     tn.IndexOf(nicName, StringComparison.OrdinalIgnoreCase) >= 0);

                if (matchesNic || tn.EndsWith(" Properties", StringComparison.OrdinalIgnoreCase))
                {
                    found = h;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static string GetText(IntPtr h)
        {
            var sb = new StringBuilder(512);
            _ = GetWindowText(h, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string GetClass(IntPtr h)
        {
            var sb = new StringBuilder(256);
            _ = GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc f, IntPtr p);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);

        private struct RECT { public int Left, Top, Right, Bottom; }
    }
}
