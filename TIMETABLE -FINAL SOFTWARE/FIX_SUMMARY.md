# Fix Summary: Preventing Elective and NML Course Overlaps

## Problem
Elective courses were being scheduled in time slots that overlapped with NML (Minor) courses for the same department and session, causing conflicts in the timetable.

## Root Cause
The elective scheduling logic (`_schedule_elective_classes` and `_schedule_elective_tutorials` methods) was designed to assign common time slots across all departments in a semester. However, it was not properly checking for conflicts with NML courses that had already been scheduled for specific departments and sessions.

The scheduling order is:
1. **Minor (NML) classes** - scheduled FIRST in early morning slots (07:30-08:30)
2. **Priority courses** (combined classes and electives) - scheduled SECOND
3. **Regular courses** - scheduled LAST

While electives were checking for general slot availability using `_is_time_slot_available_global()`, they were not specifically checking if those slots conflicted with NML courses across all departments in the semester.

## Solution
Modified both `_schedule_elective_classes` and `_schedule_elective_tutorials` methods in `schedule_generator.py` to add an additional conflict check:

### Key Changes:
1. **Added NML conflict detection**: Before assigning a time slot to electives, the code now checks if that slot conflicts with NML courses across ALL departments in the semester (both Pre-Mid and Post-Mid sessions).

2. **Implementation**: The fix iterates through all departments and sessions to check if the proposed elective slot is occupied by a Minor course (identified by checking if the slot is in `MINOR_SLOTS`).

3. **Skip conflicting slots**: If a conflict with NML courses is detected, that slot is removed from consideration and the algorithm tries a different slot.

### Code Logic:
```python
# Check if this slot conflicts with NML courses across ALL departments
conflicts_with_nml = False
for dept in DEPARTMENTS:
    for sess in [PRE_MID, POST_MID]:
        # Check if slot is occupied by NML courses
        if not self._is_time_slot_available_global(day, slots, dept, sess, semester_id):
            # Check if the conflict is with a Minor course
            if slot in MINOR_SLOTS:
                conflicts_with_nml = True
                break

if conflicts_with_nml:
    # Skip this slot and try another
    continue
```

## Verification
After implementing the fix:
- ✅ Timetable generation completed successfully
- ✅ No room allocation conflicts detected
- ✅ Electives are now scheduled in slots that don't overlap with NML courses
- ✅ All departments maintain proper separation between NML and elective courses

## Files Modified
- `schedule_generator.py`:
  - `_schedule_elective_classes()` method (lines ~948-1000)
  - `_schedule_elective_tutorials()` method (lines ~1102-1154)

## Impact
This fix ensures that:
1. Students can attend both NML courses and electives without conflicts
2. The timetable respects the scheduling constraints for all course types
3. Electives maintain their common slots across departments while avoiding NML conflicts
