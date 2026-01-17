﻿namespace NetIsolatePlus.Models
{
    public class NicInfo
    {
        // Stable identity: prefer GUID (NetCfgInstanceId), then DeviceID, then PNPDeviceID.
        // Never rely on friendly names as identity (users can rename adapters).
        public string Id { get; init; } = "";

        // Raw identity fields (kept for diagnostics + future matching like Status-by-GUID)
        public string Guid { get; init; } = "";          // Win32_NetworkAdapter.GUID
        public string DeviceId { get; init; } = "";      // Win32_NetworkAdapter.DeviceID
        public string PnpDeviceId { get; init; } = "";   // Win32_NetworkAdapter.PNPDeviceID

        // Display fields
        public string Name { get; init; } = "";          // Friendly name (NetConnectionID)
        public string Description { get; init; } = "";   // Adapter description

        // Admin state (enabled/disabled). This is NOT "connected".
        public bool Enabled { get; init; }

        // Extra state we will use in later phases (no UI changes yet)
        public bool? NetEnabledRaw { get; init; }        // Win32_NetworkAdapter.NetEnabled (can be null/unreliable)
        public int? NetConnectionStatus { get; init; }   // Win32_NetworkAdapter.NetConnectionStatus (link state)
        public uint? ConfigManagerErrorCode { get; init; } // 22 = disabled
        public bool? PhysicalAdapter { get; init; }      // Win32_NetworkAdapter.PhysicalAdapter
    }
}
