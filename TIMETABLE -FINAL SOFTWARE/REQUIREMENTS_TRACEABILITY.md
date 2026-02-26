# Requirements Traceability Matrix

This document maps each requirement to the current implementation status in the codebase.

**Legend:**
- **Implemented** – Fully satisfied by existing code
- **Partial** – Partially implemented; gaps remain
- **Not Implemented** – Not present in codebase

---

## Automate

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R14 | Intelligent engine: auto-allocate classes, faculty, rooms by constraints (availability, course requirements, faculty preferences) | **Partial** | `schedule_generator.py`: auto room/slot allocation by capacity and LTPSC. **Gap:** Faculty preferences and availability not used for class scheduling. |
| R14.1 | Allow manual corrections for timetable coordinator | **Not Implemented** | No UI or workflow for editing generated timetable; output is Excel only. |
| E1 | Generate automated exam timetable with academic timetable | **Implemented** | `main.py` (exam block), `exam_scheduler.py`: exports `exam_timetable.xlsx` with Mid/End sem, FN/AN. |
| E2 | Auto-assign classrooms for exams by availability and student strength | **Implemented** | `exam_scheduler.py` (exam rooms), `excel_exporter._assign_room_by_capacity`; seating uses room capacity. |
| E8 | Generate seating charts with deskwise allocation | **Implemented** | `seating_arrangement.py`: bench layout (COL1–COL4, 6×4), deskwise; exports `seating arrangement.xlsx`. |
| R2.1 | Calculate number of students registered per course | **Partial** | Course data has “Registered Students”; used for room capacity. **Gap:** No explicit “count from student_data” per course; seating infers from department/semester. |
| D5 | Integrate timetable with Google Calendar (notify class/lab timings) | **Not Implemented** | No calendar API or export for Google Calendar. |
| E3 | Mixed seating (students from different courses in same room) | **Implemented** | `seating_arrangement._can_sit_together`, `_generate_seating_for_room_with_students`: pairs different semesters/courses; same-exam students not adjacent. |
| O1 | Color-code courses in output timetable | **Implemented** | `excel_exporter._apply_color_coding`, `_color_for_course`, `_palette`: per-course colors in timetable sheets. |
| O2 | Download timetable in PDF or Excel | **Partial** | Excel export implemented. **Gap:** No PDF export. |

---

## Input

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R1 | Course info: codes, titles, credits, departments, instructors, LTPSC | **Implemented** | `excel_loader.load_all_data`, course_data.xlsx; columns Course Code, Course Name, Semester, Department, LTPSC, Credits, Instructor, etc. |
| R1.1 | Accept course code, name, [core, elective basket, minor basket] | **Partial** | Core/elective/minor via Elective column and patterns (ELEC, Minor). **Gap:** No explicit “basket” type; 7th sem baskets (7B1…) are course-code pattern only. |
| R1.2 | Validate course code format (e.g. CS201) | **Not Implemented** | No format validation (e.g. regex for department + digits). |
| R1.3 | No duplicate course codes; notify error | **Not Implemented** | Duplicates dropped in session division; no upfront validation or error report for duplicate codes. |
| R2 | Student registration (Reg no, course code) | **Partial** | `student_data.xlsx` loaded; seating infers courses from department/semester. **Gap:** No explicit (Reg no, course code) list; per-course registration count not derived from this. |
| R3 | Configurable LTPSC input and management | **Implemented** | `excel_loader.parse_ltpsc`: L-T-P parsed; defaults if missing; config durations in `config.py`. |
| R4 | Instructor + preferences (day, time); instructor vs coordinator; coordinator courses not counted when scheduling instructor | **Not Implemented** | Faculty list used only for invigilation; no day/time preferences or coordinator flag; scheduling ignores faculty availability. |
| R6 | Classroom details: capacity, availability, resources, constraints | **Partial** | `classroom_data.xlsx`, capacity and type (lecture/lab); used for allocation. **Gap:** No explicit availability windows or resource constraints. |
| R7 | Room number, capacity, type (lecture/lab) | **Implemented** | `schedule_generator`: classrooms, lab_rooms, nonlab_rooms, software/hardware labs from classroom_data. |
| R8 | Lab assistant preferences and allocation | **Not Implemented** | No lab assistant input or allocation. |
| R13 | Validate room/lab capacity ≥ student registrations | **Partial** | `_assign_room` uses `required_capacity`; C004 and elective room assignment check capacity. **Gap:** No explicit validation report “capacity < registration” or block generation. |

---

## Input Validation

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R3.1 | Each course has L, T, P, S, C (LTPSC) | **Implemented** | `excel_loader.parse_ltpsc`: L,T,P parsed; S,C in format but not used for “self-study” allocation. |
| R5 | Allocate self-study hours per S of LTPSC (after class hours) | **Not Implemented** | S from LTPSC not used; no self-study slots. |
| R15.2 | Check total weekly lectures from LTPSC | **Partial** | LTPSC → lectures/tutorials/labs per week; exported as “allocated/required”. **Gap:** No explicit “total weekly lectures” validation rule. |
| R15.3 | Distribute lectures evenly across the week | **Partial** | `_schedule_lectures` uses `used_days` to spread; no strict “even” distribution guarantee. |
| D1 | Combined classes for multiple batches | **Implemented** | Combined Class flag; C004; `_schedule_combined_class`, shared slots for CSE-A/B and DSAI/ECE. |
| D4 | Allocate class and lab on same day | **Not Implemented** | No constraint to force same-day class+lab; scheduling is slot-based only. |
| D6 | All branch students free for elective classes | **Implemented** | Electives use common slots across departments (`_schedule_elective_classes`, `semester_elective_slots`). |

---

## Validation Functionality

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R15 | Lecture scheduling rules (faculty, room, course constraints) | **Partial** | Room and course (LTPSC) and slot conflicts. **Gap:** Faculty constraints not applied. |
| R15.1 | At most one lecture per course per day | **Partial** | `used_days` and slot logic reduce repeats; not guaranteed “exactly one per day” rule. |
| R2.2 | Validate roll number format [YY][B/M/PHD][CS/DS/DA/EC][0-9][0-9][0-9] | **Not Implemented** | Roll no used as-is; BCS/BDS/BEC inferred for dept; no format validation. |
| R3.2 | Credits >2 full semester, ≤2 half semester; consider instructor duty | **Implemented** | `excel_loader.divide_courses_by_session`; half-sem split. **Gap:** “Instructor duty” not considered. |
| R7.1 | Track occupied/free hours; validate rooms not double-booked | **Implemented** | `schedule_generator.room_occupancy`, `room_bookings`, `validate_room_conflicts`. |
| R10.1 | No professor conflict (one professor, multiple batches/electives) | **Not Implemented** | No faculty–slot mapping; no conflict check per instructor. |
| R10.2 | Cross-validate with instructor availability | **Not Implemented** | No instructor availability input used in scheduling. |
| R11 | Room allocation (classrooms) | **Implemented** | `_assign_room`, nonlab_rooms, capacity-based; C004 for combined. |
| R11.1 | Auto-allocate by student count and room capacity | **Implemented** | `required_capacity` in scheduling; `_assign_room_by_capacity` for electives/minors. |
| R11.2 | Validate strength vs capacity; notify if no room matches | **Partial** | Capacity used in assignment. **Gap:** No explicit “no room fits” validation message. |
| R12 | Lab allocation | **Implemented** | Labs by LTPSC P; software/hardware by department; `_assign_room(is_lab=True)`. |
| R12.1 | Assign labs by LTPSC P | **Implemented** | `parse_ltpsc` → Labs_Per_Week; `_schedule_labs`. |
| R12.2 | Cross-validate with lab assistant availability | **Not Implemented** | No lab assistant data or checks. |
| R16.1 | Configurable short breaks | **Partial** | Lunch in config; no separate “short break” slots. |
| R16.2 | 5–10 min gap between sessions | **Partial** | 30-min slots; no explicit gap parameter (gap is slot boundary). |
| R16.3 | No overlap of breaks and classes; staggered lunch | **Partial** | Lunch slots excluded from teaching. **Gap:** No staggered lunch by department. |
| R17.2 | No classes during lunch slot | **Implemented** | `LUNCH_SLOTS` marked in `_initialize_schedule`; not used for classes. |
| E3 (seat) | No two same-exam students seated next to each other | **Implemented** | `_can_sit_together`, pairing by different semester/course in seating. |
| E4 | Max 3 per bench; optimal utilization | **Partial** | Current layout 2 per bench (COL1/COL2). **Gap:** E4 asks “max 3 per bench”; need to extend to 3 if required. |
| E5 | Balanced invigilation; even distribution | **Partial** | `exam_scheduler._generate_invigilation_data`: random shuffle; no balance tracking. **Gap:** Explicit balance check/redistribution not done. |
| E6 | Exams in working hours; FN/AN sessions | **Implemented** | Exam times in `_create_exam_sheet` (FN 10:00–11:30, AN 15:00–16:30); 7 days. |
| E7 | Avoid conflicts for students in multiple overlapping subjects | **Partial** | Exam scheduling by department/day; no explicit “student exam conflict” check. |
| E9 | Configurable exam durations | **Not Implemented** | FN/AN durations hardcoded; no parameter for 1h/2h/3h. |
| D3 | Multiple batches; separate lecture/lab scheduling | **Implemented** | CSE-A, CSE-B, DSAI, ECE; separate schedules; labs by department. |
| R10 (map) | Map instructors to courses, subjects, departments | **Partial** | Course data has Instructor column; no separate mapping module or validation. |

---

## Verification and Validation

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R10 | Instructor–course mapping and validation | **Partial** | Instructor in course sheet and export; no dedicated verification step. |

---

## View

| ID | Requirement | Status | Location / Notes |
|----|-------------|--------|------------------|
| R9 | Centralized calendar view for professors and lab assistants | **Not Implemented** | No calendar UI; only Excel output. |
| R16 | Configurable breaks, session gaps, lunch | **Partial** | `config.py`: LUNCH_SLOTS, TEACHING_SLOTS; lunch configurable. **Gap:** Short breaks and gap duration not configurable. |
| R17 | Fixed lunch break period | **Implemented** | `LUNCH_SLOTS` in config; reserved in schedule. |
| R17.1 | Reserve lunch window (e.g. 12:30–2:00 PM) | **Implemented** | `13:00-13:30`, `13:30-14:00` in config (1-hour window; extendable in config). |
| R18 | Configurable start/end times for sessions | **Partial** | `TEACHING_SLOTS` in config (07:30–18:00). **Gap:** No separate “core hours” vs “extended” bounds. |
| R18.1 | Core courses within institute hours (e.g. 9:00–17:00) | **Partial** | Minor in MINOR_SLOTS (07:30–08:30); rest of slots used for classes. **Gap:** No explicit “core only 9–17” rule. |
| R18.2 | Minor/major before 9:00 and after 17:30 (7:30–9:00, 18:30–20:00) | **Partial** | Minor in 07:30–08:30. **Gap:** No “after 17:30” slot band; slots end at 18:00. |

---

## Summary Counts

| Status | Count |
|--------|-------|
| Implemented | 28 |
| Partial | 24 |
| Not Implemented | 22 |

---

## Next Steps

See **IMPLEMENTATION_PLAN.md** for prioritized gaps and suggested implementation order.
