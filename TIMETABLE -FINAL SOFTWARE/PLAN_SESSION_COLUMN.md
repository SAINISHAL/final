# Plan: Add "Pre-Mid / Post-Mid / Full" Column to Course Data

**Status: Implemented.** Add one column in `course_data.xlsx` so the coordinator can explicitly set whether each course runs in **Pre-Mid**, **Post-Mid**, or **Full** (both sessions).

---

## 1. Current Behaviour (No New Column)

Session assignment is **derived** in `excel_loader.divide_courses_by_session()`:

- **Electives** → always **Full** (both Pre-Mid and Post-Mid)
- **HSS** → **Full**
- **Credits > 2** → **Full**
- **Credits ≤ 2** → **half semester**: split by Course Code (first half → Pre-Mid, second half → Post-Mid)
- **CSE-A and CSE-B** → same course set in both sessions (shared half-sem list)

Downstream:

- `schedule_generator.generate_department_schedule()` uses `divide_courses_by_session()` to get Pre-Mid / Post-Mid course lists for timetables.
- `excel_exporter._get_course_details_for_session()` uses the same for course details.
- `exam_scheduler.get_all_pre_mid_courses()` and `get_all_post_mid_courses()` use it for Mid-sem and End-sem exam lists.

So **one place** drives session assignment: `divide_courses_by_session()` in `excel_loader.py`.

---

## 2. Proposed Input Change

**New column in course_data (e.g. `Session` or `Course Session`):**

| Value    | Meaning                          |
|----------|----------------------------------|
| Pre-Mid  | Course runs **only** in Pre-Mid  |
| Post-Mid | Course runs **only** in Post-Mid |
| Full     | Course runs in **both** sessions |

- Allow case-insensitive, trimmed values; accept variants (e.g. "Pre", "Pre-Mid", "Post", "Post-Mid", "Full", "Both").
- **Blank / missing** → keep current derived behaviour (electives, credits, HSS, half-sem split).

**Override rule:** If the column has a value for a row, that value **overrides** the current rules for that course (e.g. an elective can be forced to "Post-Mid" only).

---

## 3. Will It Work? (Feasibility)

**Yes.** The design is consistent and fits the existing flow:

1. **Single point of change:** Only `divide_courses_by_session()` needs to look at the new column; all callers (schedule_generator, excel_exporter, exam_scheduler) already use its return values, so they will automatically respect Pre-Mid / Post-Mid / Full once the loader is updated.
2. **Backward compatible:** If the column is missing or blank, behaviour stays as today (derived from Elective, Credits, HSS).
3. **2-credit sharing:** `_apply_two_credit_sharing()` already avoids moving electives. We must also avoid moving courses that have **explicit** Pre-Mid or Post-Mid (so DSAI/ECE alignment does not override the coordinator’s choice). That only needs an extra “do not move” set passed into `_apply_two_credit_sharing()`.

---

## 4. Files to Change (When You Implement)

| File                 | Change |
|----------------------|--------|
| **excel_loader.py**  | **Main logic.** In `divide_courses_by_session()`: (1) Detect column (e.g. "Session", "Course Session"). (2) Normalize values to Pre-Mid / Post-Mid / Full. (3) Partition rows: explicit Pre-Mid only, Post-Mid only, Full, and rest (blank). (4) For **rest**, keep current logic (electives → Full, HSS → Full, credits>2 → Full, credits≤2 → split). (5) Build `pre_mid_courses` = explicit Pre-Mid + explicit Full + derived Pre-Mid; `post_mid_courses` = explicit Post-Mid + explicit Full + derived Post-Mid. (6) When calling `_apply_two_credit_sharing()`, pass a set of course codes that have explicit Pre-Mid or Post-Mid so they are **not** moved. |
| **excel_loader.py**  | In `_apply_two_credit_sharing()`: add optional parameter `explicit_session_codes=None`. Skip moving any course whose code is in this set (same as electives). |
| **config.py**       | Optional: add constants for column name(s) and allowed values (e.g. `SESSION_COLUMN_NAMES`, `SESSION_VALUES`) for validation/docs. |
| **exam_scheduler.py** | No change (uses `divide_courses_by_session()`). |
| **schedule_generator.py** | No change (uses `divide_courses_by_session()`). |
| **excel_exporter.py**   | No change (uses same division via schedule_gen / ExcelLoader). |
| **README / REQUIREMENTS_TRACEABILITY** | Document the new column and values. |

---

## 5. Edge Cases to Handle in Code (Later)

- **Invalid value** (e.g. "Mid", "Unknown"): treat as blank and fall back to current logic; optionally log a warning.
- **Elective marked "Post-Mid" only:** override; course appears only in Post-Mid (coordinator’s choice).
- **CSE-A and CSE-B:** If both have the same course with "Pre-Mid", both get it in Pre-Mid only; no change to “same course set” for CSE-A/B.
- **2-credit shared course with explicit "Pre-Mid":** Do not move it to Post-Mid for DSAI/ECE; respect explicit session (hence passing `explicit_session_codes` into `_apply_two_credit_sharing`).

---

## 6. Summary

| Question | Answer |
|----------|--------|
| Is adding a Pre-Mid / Post-Mid / Full column OK? | **Yes.** |
| Will it work with the rest of the system? | **Yes;** only `excel_loader.py` (and optionally `config.py`) need changes; timetable, export, and exam logic stay as-is. |
| Backward compatible? | **Yes;** blank/missing column → current derived behaviour. |
| 2-credit sharing conflict? | **No;** we skip moving courses that have explicit Pre-Mid or Post-Mid. |

---

## 7. How to Use (course_data.xlsx)

Add a column with one of these names (first found is used):

- **Session**
- **Course Session**
- **Pre/Post/Full**

Allowed values (case-insensitive):

| Value in Excel | Meaning |
|----------------|--------|
| **Pre** or **Pre-Mid** | Course runs **only** in Pre-Mid session |
| **Post** or **Post-Mid** | Course runs **only** in Post-Mid session |
| **Full** or **Both** | Course runs in **both** Pre-Mid and Post-Mid sessions |
| *(blank)* | Use derived logic (electives/credits/HSS/half-sem split) |

If the column is missing or a row is blank, behaviour is unchanged (electives → Full, credits > 2 → Full, credits ≤ 2 → split).
