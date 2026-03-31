# 🛠️ DeviceMan: Ultimate Hardware Auditor & Diagnostics

**DeviceMan** is a professional-grade, open-source alternative to paid diagnostic tools. Developed in VB.NET with a focus on performance and security, it provides deep system auditing and critical driver maintenance utilities.

## 🌟 Key Features (State-of-the-Art)

### 1. Secure Driver Backup & Restoration
This is the core strength of DeviceMan. It doesn't just copy files; it ensures system stability:
* **X509 Certificate Validation:** Automatically verifies digital signatures in `.cat` files to ensure driver integrity before restoration.
* **Platform Awareness:** Prevents BSODs by blocking cross-architecture driver injection (x86/x64 mismatch).
* **Audit-Ready Logs:** Generates real-time console logs and professional HTML5/Bootstrap reports for post-maintenance verification.

### 2. High-Performance System Monitor (SysRes)
A "State-of-the-Art" UI built with optimized GDI+ rendering:
* **Real-time Graphics:** Per-thread CPU history charts and GPU load/temp sensors via OpenHardwareMonitor[cite: 139, 142].
* **Low Overhead:** Uses smart WMI caching and PerformanceCounters to keep resource usage minimal[cite: 131, 133].
* **Compact Overlay Mode:** A semi-transparent "Always-on-Top" dashboard for monitoring during heavy workloads.

### 3. Advanced Task Manager (TskMgr)
* **Process Traceability:** Real-time PID tracking, memory allocation, and owner identification[cite: 162, 166].
* **Safety First:** Prevents accidental termination of critical system processes or the auditor itself.

## 🚀 Getting Started

### Prerequisites
* Windows 10/11
* .NET Framework 4.8 or .NET 6/7/8
* **Administrator Privileges:** Required for sensor data collection and driver injection via `pnputil`.

## 🤝 Contributing
DeviceMan is modular and community-driven. We welcome contributions in:
* Adding new WMI sensor modules.
* Improving driver backup compression algorithms.
* Localizing the UI for more languages.

---
Developed by **Jonni** | Part of the open-source hardware initiative.
