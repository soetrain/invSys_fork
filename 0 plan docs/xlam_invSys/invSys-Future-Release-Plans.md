# invSys Future Release Plans (R2+)
Extracted from invSys-Design-v4.md to keep Release 1 authoritative.

## Release 2/3: F# and Rust Tooling (Future)
This document captures the post-R1 roadmap for introducing compiled tooling. Nothing in this file is required for Release 1.

### Summary Table
| Component | R1 (VBA) | R2 (F#) | R3 (Rust) |
|-----------|----------|---------|-----------|
| HQ Aggregation | VBA macro (30-60s) | F# CLI (5-10s) | Rust CLI (2-5s) |
| Backup | Manual or simple VBA | F# CLI with rotation | Rust Windows Service |
| Schema Validation | VBA self-repair | F# CI/CD validator | Rust build-time checks |
| Item Search | VBA local cache | F# HTTP API + cache | Rust high-throughput API |
| Monitoring | Manual Admin UI | F# scheduled checks | Rust embedded service |

## Component Plans (R2/R3)
### HQ Aggregation
- R2: F# CLI using ClosedXML. Reads warehouse snapshots without Excel automation. Writes `Global.InventorySnapshot.xlsb`.
- R2 runtime: Task Scheduler launches a standalone EXE; structured logs.
- R3: Rust port (calamine + static binary) for faster startup and lower memory.

### Backup
- R2: F# CLI that runs on schedule, performs timestamped backups, and verifies checksums.
- R3: Rust service for always-on backup/restore automation with low overhead.

### Schema Validation
- R2: F# CLI validator for CI/CD and pre-deploy checks against schema manifests.
- R3: Rust validator for build-time enforcement and embedded checks.

### Item Search
- R2: F# HTTP API + cache for real-time search, Excel can fall back to local cache.
- R3: Rust API for low-latency, high-throughput search at scale.

### Monitoring
- R2: F# scheduled health checks with email alerts.
- R3: Rust service for always-on monitoring and low resource usage.

## F# as Executable Specification
All F# implementations serve as verified specifications for eventual Rust ports. Key benefits:
1. Type safety: compile-time guarantees vs dynamic typing.
2. Token efficiency: concise functional pipelines for AI prompts.
3. Syntactic similarity: F# pattern matching maps cleanly to Rust enums.
4. Testability: FsCheck property tests validate business rules.
5. Bridge language: research suggests F# -> Rust translation is stronger than many direct pairs.

## When to Port to Rust
Wait until:
- F# version is stable and proven in production (6+ months).
- Performance profiling identifies real bottlenecks.
- Scale justifies investment (>10 warehouses, >100k SKUs, >1000 req/sec).

Rust gains are most visible when:
- High-volume I/O (HQ aggregating 10+ warehouses).
- Low-latency requirements (item search <1ms).
- Embedded deployment (Windows Service, static binary).
- Cross-platform needs (Linux servers for CI/CD).

## Research Notes (2025-2026)
- LANTERN: cross-language translation to Rust via functional intermediates improved repair success by 13-30%.
- XL-CoGen: bridging languages for Rust translation outperformed direct translation.
- Claude 3.5: 43.5% Pass@1 for real-world C++ -> Rust (functional intermediates help).

## Conclusion
Release 1 stays VBA-only with zero external runtime dependencies. Future evolution:
1. Replace VBA back-end components with F# where performance matters.
2. Use F# as executable specifications and verified reference implementations.
3. Port select F# services to Rust for high-scale, low-latency scenarios.
4. Maintain F# test suites as oracle for Rust implementation correctness.