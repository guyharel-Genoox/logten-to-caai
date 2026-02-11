# CAAI Flight Time Classification Rules

This document explains the Israeli Civil Aviation Authority (CAAI / רת"א) rules for classifying flight time on the tofes-shaot (טופס שעות) form.

## Aircraft Groups

| Group | Hebrew | Description | Examples |
|-------|--------|-------------|----------|
| A (א) | בוכנה חד מנועי | Single-engine piston | C172, C150, PA28, SR22 |
| B (ב) | בוכנה רב מנועי | Multi-engine piston | PA44, BE76 |
| C (ג) | סילון/טורבו פרופ רב מנועי | Multi-engine jet/turboprop | A319, A320, H25B |
| D (ד) | טורבו פרופ חד מנועי | Single-engine turboprop | - |

## Role Classification Rules

### Rule 1: Student (מתלמד)
- Time logged **only** during dual instruction (with instructor)
- If an instructor is present OR dual_received > 0, the flight is classified as Student
- Student time cannot also be PIC time

### Rule 2: PIC (טייס מפקד)
PIC time is logged when you are the pilot in command, **excluding**:
- Flights with an instructor present (those are Student)
- Safety pilot time on single-engine aircraft
- SIC time on multi-engine aircraft

### Rule 3: Safety Pilot on Single-Engine
- A safety pilot on a single-engine aircraft is **NOT** PIC (Rule 4 of the form)
- A safety pilot on a single-engine aircraft is **NOT** SIC (no SIC concept on SE)
- These hours are **EXCLUDED** from the form category totals entirely
- The form row total = PIC + SIC + Student (safety pilot hours not counted)

### Rule 4: SIC (טייס משנה)
- SIC time only exists on multi-engine aircraft requiring two pilots
- On single-engine aircraft, if both PIC and SIC fields have values, it's treated as PIC
- SIC half-credit per regulation 42(b): Grand total = PIC + SIC/2 + Student

### Rule 5: PIC Cross-Country
- PIC XC hours exclude safety pilot time and instructor flights
- Cross-country threshold: > 50km (~27 nautical miles) from departure

## Table Structure

### Table 1: Aircraft Flight Hours (סיכום שעות טיסה)
- One row per aircraft type
- Category column (A/B/C/D) shows the form_total (PIC + SIC + Student for that type)
- Day columns: PIC, PIC XC, SIC, Student
- Night columns: PIC, PIC XC, SIC, Student
- **Simulators are NOT included** in Table 1

### Table 2: Instrument Time
- Actual instrument time in aircraft
- Simulated instrument time in aircraft
- Simulator device time (separate column)

## Instrument Time Rules

### Rule 6: Actual Instrument (not during instruction)
On single-pilot aircraft, actual instrument time when **not** under instruction counts as PIC time.

### Rule 7: Simulator Time
Simulator/training device time is **never** included in total flight hours.
It appears only in the instrument time table.

### Rule 8: Simulated Instrument in Air
Simulated instrument time during dual instruction is classified as **Student** time (not PIC).

## SIC Half-Credit (תקנה 42(ב))

Per Israeli aviation regulation 42(b):
- Grand total = PIC + (SIC / 2) + Student
- This applies to the overall total on the form (Row 47, Column O)
- Individual columns show full PIC, full SIC, and full Student hours
- The formula applies the half-credit when calculating the grand total

## CPL Requirements (רישיון טיס מסחרי)

| Field | Description |
|-------|-------------|
| C12 | PIC XC time (excludes safety pilot and instructor flights) |
| C13 | Total dual received |
| C14 | Dual instrument instruction (actual + simulated, during instruction only) |
| C15 | Night landing count |
| C16 | Night flight hours |
| C17 | Longest solo XC flight (hours, date, distance in km, route) |
| C18 | Complex time: retractable gear + variable pitch prop, OR Group B, OR Group C |

## ATPL Requirements (רישיון טיס תובלה)

| Field | Description |
|-------|-------------|
| C13 | Total XC navigation hours (all roles) |
| C14 | Night PIC XC hours |
| C15 | Total instrument time (actual + simulated in aircraft, NOT device) |
