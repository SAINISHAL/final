# Implementation Plan – Requirements Gaps

This document lists requirements that are **Partial** or **Not Implemented** and suggests implementation steps. Items are grouped by theme and priority.

---

## Priority 1 – Input Validation & Data Integrity

These affect correctness of data before scheduling and are low-effort.

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **R1.2** | Validate course code format (e.g. CS201) | In `excel_loader.load_all_data` (or a new `validators.py`): after loading course_data, run regex per row (e.g. `^[A-Z]{2,4}\d{3}$`). Optionally make pattern configurable in config. |
| **R1.3** | No duplicate course codes; notify error | In `excel_loader.load_all_data`: `course_df['Course Code'].duplicated()`; collect duplicates, print/return error list and optionally raise or flag so generation does not proceed. |
| **R2.2** | Validate roll number format [YY][B/M/PHD][CS/DS/DA/EC][0-9][0-9][0-9] | In `excel_loader` or `validators.py`: when loading student_data, validate Roll No with regex; report invalid rows. |
| **R13** | Validate room/lab capacity ≥ student registrations | After room assignment in `schedule_generator`: for each (course, room), compare room capacity with course “Registered Students”; log/collect violations and optionally add a “validation summary” sheet or console report. |
| **R11.2** | Notify if no room matches student strength | In `_assign_room`: when no room has capacity ≥ required_capacity, append to a “allocation_failures” list; at end of export, report these courses (e.g. in Course_Summary or a Validation sheet). |

---

## Priority 2 – Faculty & Lab Assistant

Required for R14 (intelligent engine) and R4/R8.

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **R4** | Instructor preferences (day, time); instructor vs coordinator; coordinator courses not counted when scheduling instructor | Extend `faculty_availability.xlsx` (or new sheet): columns e.g. Faculty, Day, TimeSlot, IsCoordinator, CourseCode. In `schedule_generator`: (1) when assigning a slot for a course, check instructor’s preferred (day, time); (2) if IsCoordinator=Yes for that course, do not block that slot for other courses of same instructor. Load in `excel_loader`; pass to ScheduleGenerator. |
| **R10.1** | No professor conflict (one professor, multiple batches/electives) | Maintain `instructor_slots[(instructor_id, day, slot)]` in ScheduleGenerator. Before placing a class, check instructor; if (instructor, day, slot) already used, skip. Requires course → instructor mapping from course_data. |
| **R10.2** | Cross-validate with instructor availability | Use same faculty availability structure as R4; before accepting a slot, check that (instructor, day, slot) is “available”. |
| **R8** | Lab assistant preferences and allocation | New input: e.g. `lab_assistant_availability.xlsx` (assistant, day, slot, lab_room). In lab scheduling, assign lab room and then assign assistant from this list; validate no double-booking (R12.2). |
| **R12.2** | Cross-validate lab assistant availability | When allocating a lab slot, pick an assistant who is available that (day, slot); if none, log and optionally leave “Unassigned” in export. |

---

## Priority 3 – Manual Correction (R14.1)

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **R14.1** | Allow manual corrections for timetable coordinator | **Option A (lightweight):** Document that coordinator can edit generated Excel (move/copy classes, change rooms) and re-run only “validation” (e.g. run room conflict + capacity checks on modified file). **Option B (full):** Add a “load existing timetable” path: read back an Excel timetable into internal structures, run validators, allow a simple CLI or script to adjust one (course, day, slot) and re-export. **Option C (UI):** Web or desktop UI to open generated timetable, edit cells, save and re-validate. Start with Option A + validation script; later Option B if needed. |

---

## Priority 4 – Output & Export

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **O2 (PDF)** | Download timetable in PDF | Add dependency (e.g. `reportlab` or `weasyprint`). New module e.g. `pdf_exporter.py`: build a simple grid (days × slots) and course names per cell from schedule DataFrame; output PDF per sheet or one PDF with multiple pages. Call from main after Excel export. |
| **R2.1 (explicit)** | Calculate students registered per course | If student_data has (Roll No, Course Code) or (Roll No, Semester) + course list per dept: aggregate count per (Course Code, Semester) and attach to course rows or a “Course Enrollment” sheet. If not, document that “Registered Students” in course_data is the source and ensure it is used everywhere (already used for capacity). |

---

## Priority 5 – Scheduling Rules & Config

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **D4** | Allocate class and lab on same day | In `_schedule_course`: when scheduling lab for a course, pass “prefer_days” = set of days already used for that course’s lectures/tutorials; in `_schedule_labs` prefer those days when choosing (day, slot). |
| **R15.1** | At most one lecture per course per day | In `_schedule_lectures`: after picking (day, start_slot), if this course already has a lecture on `day`, skip (enforce one-lecture-per-day). |
| **R5** | Allocate self-study hours per S of LTPSC | Parse S from LTPSC; add “self-study” slots (e.g. after 17:30 or in a separate “self-study” view). Optional: add a row in timetable “Self-Study (Course X)” in non-teaching slots; or just report “Self-Study hours” in course details. |
| **R15.3** | Distribute lectures evenly across week | In `_schedule_lectures`: sort candidate (day, slot) by “current load on day” (e.g. count of slots already used that day for this dept/session); prefer days with lower load so distribution is more even. |
| **R16.1 / R16.2** | Configurable short breaks; 5–10 min gap | In config: e.g. `SHORT_BREAK_SLOTS = []` and `MIN_GAP_SLOTS = 1` (1 × 30 min = “gap”). When marking slots busy, reserve MIN_GAP_SLOTS after each class end; optionally reserve SHORT_BREAK_SLOTS as non-teaching. |
| **R16.3** | Staggered lunch by department | In config: e.g. `LUNCH_BY_DEPT = {'CSE-A': ['12:30-13:00', ...], 'DSAI': ['13:00-13:30', ...]}`. In schedule, mark lunch per department instead of global LUNCH_SLOTS. |
| **R18.1 / R18.2** | Core 9–17; minor/major before 9 and after 17:30 | In config: `CORE_SLOTS` (9:00–17:00), `MINOR_SLOTS` (7:30–9:00), `EXTENDED_SLOTS` (17:30–20:00). In scheduling: core courses only from CORE_SLOTS; minor already in MINOR_SLOTS; optionally use EXTENDED_SLOTS for “major” or extra classes. Extend TEACHING_SLOTS to 20:00 if needed. |
| **E9** | Configurable exam durations | In config: `EXAM_FN_DURATION_MIN`, `EXAM_AN_DURATION_MIN`. In `exam_scheduler._create_exam_sheet` use these for display; if you add per-exam duration later, store in exam_data and use here. |

---

## Priority 6 – Seating & Invigilation

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **E4** | Max 3 per bench; optimal utilization | In `seating_arrangement._generate_seating_for_room_with_students`: change layout to 3 students per bench (e.g. COL1/COL2/COL3); adjust max_seats to 6×4×3 or similar; keep “no same exam adjacent” logic (row/column neighbours). |
| **E5** | Balanced invigilation | In `exam_scheduler._generate_invigilation_data`: maintain “invigilation_count[faculty]”; when assigning, choose faculty with minimum count first (round-robin or min-count) instead of pure random. |
| **E7** | Avoid student exam conflicts | When building exam schedule (FN/AN per day), for each student derive “exams per day”; if a student has two exams same day (e.g. two departments), move one to another day or session and re-check. Requires student → course list from student_data. |

---

## Priority 7 – Views & Integration

| Req | Description | Suggested Implementation |
|-----|-------------|---------------------------|
| **R9** | Centralized calendar view for professors and lab assistants | Requires UI: e.g. web app (Flask/FastAPI + HTML/JS) or desktop (Electron/Tk). Load generated timetables + faculty list; show week view (days × slots) and filter by professor or lab assistant. Out of scope for current CLI/Excel-only project unless you add a minimal UI. |
| **D5** | Google Calendar integration | After export: read generated Excel, build events (course, day, time, room); use Google Calendar API (credentials, OAuth) to create events. New script e.g. `export_to_google_calendar.py`; document setup (service account or OAuth). |
| **R1.1 (basket)** | Explicit core / elective basket / minor basket | In course_data add column “Type” or “Basket” (Core / Elective Basket / Minor Basket). In loader, map to existing elective/minor logic and 7th-sem basket codes; use for validation and labels in export. |

---

## Suggested Order of Work

1. **Phase 1 (validation):** R1.2, R1.3, R2.2, R13, R11.2 – validators + report.
2. **Phase 2 (faculty):** R4, R10.1, R10.2 – faculty availability and conflict checks.
3. **Phase 3 (output):** O2 (PDF), R14.1 (manual edit doc + optional “load and validate” script).
4. **Phase 4 (scheduling tweaks):** D4, R15.1, R15.3, R16.x, R18.x, E9.
5. **Phase 5 (seating & exams):** E4, E5, E7; then R8, R12.2.
6. **Phase 6 (optional):** R9, D5, R1.1 basket type.

Use **REQUIREMENTS_TRACEABILITY.md** to mark items as Implemented/Partial as you complete them.
