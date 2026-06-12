# EIT Multiplexer PCB

EE499 Capstone Project — Low-power analog multiplexer board for Electrical Impedance Tomography (EIT).

## Overview

This board drives a 16-channel analog multiplexer (ADG1606) controlled by an RP2040 microcontroller. It is designed for EIT data acquisition, where the multiplexer selects electrode channels to inject current and measure voltage.

**Key ICs:**
- Raspberry Pi RP2040 — microcontroller
- ADG1606 — 16:1 analog multiplexer
- AP2112K — 5V to 3.3V LDO regulator
- LM27761 — negative voltage converter (for symmetric supply)

## Repository Structure

```
Hardware/         — KiCad project files (.kicad_pro, .kicad_sch, .kicad_pcb)
                    Custom footprint and symbol libraries bundled for portability
Production/
  Gerbers/        — Gerber files + drill files ready for PCB fabrication
  BOM/            — Bill of Materials (CSV and XLSX)
Docs/             — Schematic PDF, assembly drawing, component datasheets
```

## Opening the Design

1. Open **KiCad** (v7 or later)
2. Open `Hardware/EIT_Multiplexer.kicad_pro`
3. All custom libraries load automatically via relative paths — no manual configuration needed

## Fabrication

Gerber files are in `Production/Gerbers/`. The board is 2-layer FR4. Upload the contents of that folder (or `EIT_Multiplexer-job.gbrjob`) to your preferred PCB fab (JLCPCB, PCBWay, OSH Park, etc.).

## License

Hardware design files licensed under [CERN Open Hardware Licence v2 – Strongly Reciprocal (CERN-OHL-S)](LICENSE).
