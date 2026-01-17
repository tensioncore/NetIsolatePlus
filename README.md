<p align="center">
  <img src="netisolateplusv1.jpg" alt="NetIsolate+ screenshot" width="600" />
</p>

# NetIsolate+

Fast NIC isolation, instant toggles, and one-click Status/Properties — built for power users and home labs.

## Download (signed build)
Signed release builds are available here:
- https://www.nickdodd.com/systems.php

## What it does
- Toggle network adapters **ON/OFF**
- **Isolate** to a single adapter (disables all others, then restore)
- Open **Status / Properties** for any adapter
- Bulk **Toggle All**
- Sort adapters (enabled/disabled first, A→Z, Z→A)
- Optional: include/exclude **virtual adapters**
- Optional: **Start with Windows** (uses Task Scheduler)

## Safety notes
- NetIsolate+ requires **Administrator** privileges (UAC prompt) to enable/disable adapters.
- If you are connected via **RDP**, disabling adapters may disconnect you. NetIsolate+ warns before risky actions.
- Adapter identity is **GUID-only** to avoid breakage from renamed adapters.

## System requirements
- Windows 10 or Windows 11
- Self-contained Release builds include the required .NET runtime
- Admin rights to perform adapter enable/disable actions

## Build from source
See: **build_readme.md**

## Changelog
See: **CHANGELOG.md**

## License
Licensed under the **Apache License 2.0** — see `LICENSE`.

## Author
Nick Dodd — https://www.nickdodd.com  
© Tensioncore Administration Services
