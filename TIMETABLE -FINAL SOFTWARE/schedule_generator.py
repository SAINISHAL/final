"""Core scheduling logic for generating timetables from Excel data."""
import pandas as pd
import random
from config import DAYS, TEACHING_SLOTS, LECTURE_DURATION, TUTORIAL_DURATION, LAB_DURATION, MINOR_DURATION
from config import PRE_MID, POST_MID, MINOR_SUBJECT, MINOR_CLASSES_PER_WEEK, DEPARTMENTS, CSE_SECTION_CAPACITY
from config import MINOR_SLOTS, LUNCH_SLOTS, FORCED_COMBINED_COURSES
from excel_loader import ExcelLoader

class ScheduleGenerator:
    """Generates weekly class schedules for semesters and departments from Excel data."""
    
    def __init__(self, data_frames):
        """Initialize ScheduleGenerator with data frames."""
        self.dfs = data_frames
        # Track global slots per semester to avoid clashes between departments in same semester
        self.semester_global_slots = {}
        # Track room occupancy per (day, slot) globally across ALL semesters
        self.room_occupancy = {}
        # Track detailed room bookings per (sem_key, day, slot) for conflict validation
        self.room_bookings = {}
        # Load classrooms (room_name, capacity) and classify by type
        self.classrooms = []
        self.lab_rooms = []
        self.software_lab_rooms = []
        self.hardware_lab_rooms = []
        self.nonlab_rooms = []
        self.c004_room = None  # Special room for combined classes only
        self.shared_rooms = set()

        try:
            cls_df = self.dfs.get('classroom')
            if cls_df is not None and not cls_df.empty:
                name_col = None
                cap_col = None
                type_col = None
                for col in cls_df.columns:
                    cl = str(col).lower()
                    if name_col is None and any(k in cl for k in ['room', 'class', 'hall', 'name']):
                        name_col = col
                    if cap_col is None and any(k in cl for k in ['cap', 'seats', 'capacity']):
                        cap_col = col
                    if type_col is None and any(k in cl for k in ['type', 'category', 'room type']):
                        type_col = col
                if name_col is None:
                    name_col = cls_df.columns[0]
                    
                # NOTE: preserve the original order from the input sheet.
                # We will not sort classrooms by capacity here so that room
                # allocation follows the order in which rooms are listed
                # in the Excel input.
                for _, row in cls_df.iterrows():
                    room_name = str(row.get(name_col, '')).strip()
                    try:
                        capacity = int(float(row.get(cap_col, 0))) if cap_col is not None else 0
                    except Exception:
                        capacity = 0
                    room_type = str(row.get(type_col, '')).strip().lower() if type_col is not None else ''
                    
                    if room_name:
                        # Special handling for C004 - combined class room only
                        room_name_upper = room_name.upper()
                        if room_name_upper == 'C004':
                            self.c004_room = (room_name, capacity)
                            self.shared_rooms.add(room_name_upper)
                            continue
                            
                        # Preserve input order
                        self.classrooms.append((room_name, capacity))
                        
                        if 'lab' in room_type:
                            self.lab_rooms.append((room_name, capacity))
                            if 'software' in room_type or 'soft' in room_type:
                                self.software_lab_rooms.append((room_name, capacity))
                            if 'hardware' in room_type or 'hard' in room_type:
                                self.hardware_lab_rooms.append((room_name, capacity))
                        else:
                            # Normal classrooms (excluding C004), in input order
                            self.nonlab_rooms.append((room_name, capacity))
                
                print(f"Room configuration loaded:")
                print(f"  - C004 (Combined): {self.c004_room if self.c004_room else 'Not found'}")
                print(f"  - Normal classrooms: {len(self.nonlab_rooms)}")
                print(f"  - Software labs: {len(self.software_lab_rooms)}")
                print(f"  - Hardware labs: {len(self.hardware_lab_rooms)}")
                
        except Exception as e:
            print(f"Error loading classroom data: {e}")
            self.classrooms = []
            self.lab_rooms = []
            self.software_lab_rooms = []
            self.hardware_lab_rooms = []
            self.nonlab_rooms = []
            self.c004_room = None
            
        # Store minor slots per semester
        self.semester_minor_slots = {}
        # Store elective slots per semester, keyed by (semester_id, elective_code)
        self.semester_elective_slots = {}
        # Store elective tutorial slots per semester, keyed by (semester_id, 'ALL_ELECTIVE_TUTORIALS') or (semester_id, elective_code, 'Tutorial')
        # This ensures all departments get elective tutorials at the same time slots for both Pre-Mid and Post-Mid
        self.semester_elective_tutorial_slots = {}
        # 240-seater combined-class capacity per semester: set of (day, slot)
        self.semester_combined_capacity = {}
        # Combined class assigned slots per course and component:
        # key=(semester_id, course_code, component['Lecture'|'Tutorial'|'Lab']) -> list[(day, slot)]
        self.semester_combined_course_slots = {}
        # Global combined course slots shared across semesters but per allowed pairing group:
        # key=('GLOBAL', group_key, course_code, component) -> list[(day, start_slot)]
        self.global_combined_course_slots = {}
        self.scheduled_slots = {}  # Track all scheduled slots by semester+department
        self.scheduled_courses = {}  # Track when each course is scheduled
        self.actual_allocations = {}  # Track actual allocated counts: key=(semester_id, dept, session, course_code), value={'lectures': X, 'tutorials': Y, 'labs': Z}
        self.assigned_rooms = {}  # Track room assignments: key=(semester_id, dept, session, course_code) -> room_name
        self.assigned_lab_rooms = {}  # Track lab room assignments: key=(semester_id, dept, session, course_code) -> room_name
        # Per-room allocation count (slot assignments) so we prefer least-used rooms and balance utilization
        self.room_allocation_count = {}

    def _initialize_schedule(self):
        """Initialize an empty schedule with Days as rows and Time Slots as columns."""
        schedule = pd.DataFrame(index=DAYS, columns=TEACHING_SLOTS)
        
        # Initialize with 'Free'
        for day in DAYS:
            for slot in TEACHING_SLOTS:
                schedule.loc[day, slot] = 'Free'
        
        # Mark lunch break (now possibly multiple 30-min slots)
        for day in DAYS:
            for lunch_slot in LUNCH_SLOTS:
                if lunch_slot in schedule.columns:
                    schedule.loc[day, lunch_slot] = 'LUNCH BREAK'
        
        return schedule
    
    def _get_consecutive_slots(self, start_slot, duration):
        """Get consecutive time slots for a given duration."""
        try:
            start_index = TEACHING_SLOTS.index(start_slot)
            end_index = start_index + duration
            if end_index <= len(TEACHING_SLOTS):
                return TEACHING_SLOTS[start_index:end_index]
        except ValueError:
            pass
        return []
    
    def _ends_at_thirty(self, slots):
        """Check if a sequence of slots ends at :30."""
        if not slots:
            return False
        last_slot = slots[-1]
        # Extract end time from slot (format: 'HH:MM-HH:MM')
        try:
            end_time = last_slot.split('-')[1]  # Get the end time
            # Check if it ends at :30
            return end_time.endswith(':30')
        except (IndexError, AttributeError):
            return False
    
    def _get_preferred_start_slots(self, duration, regular_slots):
        """Get start slots that result in courses ending at :30.
        Returns (preferred_slots, remaining_slots)."""
        preferred = []
        remaining = []
        
        for start_slot in regular_slots:
            slots = self._get_consecutive_slots(start_slot, duration)
            # Check all slots are in regular_slots and not in excluded slots
            if len(slots) == duration and all(s in regular_slots for s in slots):
                if self._ends_at_thirty(slots):
                    preferred.append(start_slot)
                else:
                    remaining.append(start_slot)
        
        return preferred, remaining
    
    def _get_dept_from_global_key(self, dept_key):
        """Extract department label from a global slot key (e.g., 'CSE-A' from 'CSE-A_Pre-Mid')."""
        return dept_key.split('_')[0] if dept_key else ''

    def _departments_can_share_slots(self, dept_a, dept_b):
        """Return True if two departments are allowed to share the same time slots."""
        if not dept_a or not dept_b:
            return False

        share_groups = [
            {"CSE-A", "CSE-B"},
        ]

        for group in share_groups:
            if dept_a in group and dept_b in group:
                return True
        return False

    def _is_time_slot_available_global(self, day, slots, department, session, semester_id):
        """Enhanced slot availability check to prevent conflicts.
        Rules:
        - Same department + same session = conflict (same students can't be in two classes)
        - Same department + different session = OK (different students)
        - Different departments = OK (different students, can share slots)
        - CSE-A and CSE-B can share slots for the same courses."""
        semester_key = f"sem_{semester_id}"
        
        # Use the same tracking system as _mark_slots_busy_global
        if semester_key not in self.semester_global_slots:
            return True  # No slots booked yet for this semester
        
        # Check for conflicts
        for slot in slots:
            # Check semester-wide conflicts
            for dept_key, used_slots in self.semester_global_slots[semester_key].items():
                if (day, slot) in used_slots:
                    # Extract department from dept_key (format: "DEPT_SESSION")
                    dept_in_slot = dept_key.split('_')[0] if '_' in dept_key else dept_key
                    session_in_slot = dept_key.split('_')[1] if '_' in dept_key else ''
                    
                    # Allow CSE-A and CSE-B to share slots (they can have same courses at same time)
                    if self._departments_can_share_slots(department, dept_in_slot):
                        continue  # Allow sharing between CSE-A and CSE-B
                    
                    # Allow same department different sessions (different students)
                    if department == dept_in_slot and session != session_in_slot:
                        continue
                    
                    # Allow different departments (different students, can share slots)
                    if department != dept_in_slot:
                        continue
                    
                    # Block: same department + same session = conflict
                    # (This means department == dept_in_slot and session == session_in_slot)
                    return False
        return True

    def _mark_slots_busy_global(self, day, slots, department, session, semester_id):
        """Mark time slots as busy in global tracker."""
        key = f"{department}_{session}"
        semester_key = f"sem_{semester_id}"
        
        if semester_key not in self.semester_global_slots:
            self.semester_global_slots[semester_key] = {}
        
        if key not in self.semester_global_slots[semester_key]:
            self.semester_global_slots[semester_key][key] = set()
        
        for slot in slots:
            self.semester_global_slots[semester_key][key].add((day, slot))
            # prepare room occupancy tracker (use (day, slot) global key consistent with _assign_room)
            occ_key = (day, slot)
            if occ_key not in self.room_occupancy:
                self.room_occupancy[occ_key] = set()
    
    def _is_slot_reserved_global(self, day, slot, semester_id, current_priority):
        """Check if a slot is reserved by a higher priority course type.
        current_priority can be: 'Elective', 'Combined', 'Minor', 'Regular', 'Lab', 'Lecture', 'Tutorial'
        Strict Priority: 1. Combined, 2. Elective, 3. Minor, 4. Lab, 5. Lecture, 6. Tutorial
        """
        semester_key = f"sem_{semester_id}"
        
        # All lower priorities must avoid Combined Classes (Priority 1)
        if current_priority != 'Combined':
            for key, assigned in self.semester_combined_course_slots.items():
                if key[0] == semester_id:
                    component = key[3]
                    duration = LECTURE_DURATION if component == 'Lecture' else (TUTORIAL_DURATION if component == 'Tutorial' else LAB_DURATION)
                    for c_day, c_start in assigned:
                        if c_day == day:
                            c_slots = self._get_consecutive_slots(c_start, duration)
                            if slot in c_slots:
                                return True

        # All lower priorities must avoid Electives (Priority 2)
        if current_priority not in ['Combined', 'Elective']:
            # Check all elective keys for this semester
            found_elective = False
            for e_key, assigned in self.semester_elective_slots.items():
                if isinstance(e_key, tuple) and e_key[0] == semester_key:
                    # Check if slot is within any assigned elective range
                    for elective_day, elective_start in assigned:
                        if elective_day == day:
                            # Use LECTURE_DURATION as it's the standard for elective slots reservation
                            elective_slots = self._get_consecutive_slots(elective_start, LECTURE_DURATION)
                            if slot in elective_slots:
                                found_elective = True
                                break
                if found_elective:
                    return True
            
            # Check elective tutorial slots
            for et_key, assigned in self.semester_elective_tutorial_slots.items():
                if isinstance(et_key, tuple) and et_key[0] == semester_key:
                    for elective_day, elective_start in assigned:
                        if elective_day == day:
                            elective_slots = self._get_consecutive_slots(elective_start, TUTORIAL_DURATION)
                            if slot in elective_slots:
                                found_elective = True
                                break
                if found_elective:
                    return True

        # All lower priorities must avoid Minor (Priority 3)
        if current_priority not in ['Combined', 'Elective', 'Minor']:
            if semester_key in self.semester_minor_slots:
                for m_day, m_start in self.semester_minor_slots[semester_key]:
                    if m_day == day:
                        m_slots = self._get_consecutive_slots(m_start, MINOR_DURATION)
                        if slot in m_slots:
                            return True

        # Regular Lectures/Tutorials must avoid Labs (Priority 4)
        if current_priority in ['Lecture', 'Tutorial', 'Regular']:
             semester_key = f"sem_{semester_id}"
             if semester_key in self.semester_global_slots:
                for dept_sess_key, used_slots in self.semester_global_slots[semester_key].items():
                    # We need a way to detect if a slot was a lab. 
                    # For now, let's rely on _is_time_slot_available_global for core conflict checks.
                    # This method is primarily for blocking across sections.
                    pass

        return False

    def _is_slot_reserved_for_electives(self, day, slot, semester_id):
        """Deprecated: Use _is_slot_reserved_global instead.
        Kept for minimal logic compatibility if needed elsewhere.
        """
        return self._is_slot_reserved_global(day, slot, semester_id, 'Combined')

    
    def _is_time_slot_available_local(self, schedule, day, slots):
        """Check if time slots are available in local schedule."""
        for slot in slots:
            if schedule.loc[day, slot] != 'Free':
                return False
        return True
    
    def _mark_slots_busy_local(self, schedule, day, slots, course_code, class_type):
        """Mark time slots as busy in local schedule."""
        suffix = ''
        if class_type == 'Lab':
            suffix = ' (Lab)'
        elif class_type == 'Tutorial':
            suffix = ' (Tut)'
        elif class_type == 'Minor':
            suffix = ' (Minor)'
        
        for slot in slots:
            current_val = schedule.loc[day, slot]
            new_val = f"{course_code}{suffix}"
            if current_val == 'Free' or current_val == '-':
                schedule.loc[day, slot] = new_val
            else:
                # Append if not already present
                if new_val not in str(current_val):
                    schedule.loc[day, slot] = f"{current_val}, {new_val}"
    
    def _log_room_booking(self, semester_id, day, slot, room_name, department, course_code, session):
        """Record a room booking so conflicts can be detected later."""
        semester_key = f"sem_{semester_id}"
        slot_key = (day, slot)
        booking = {
            'room': room_name,
            'dept': department,
            'course': str(course_code).strip(),
            'session': session
        }
        if semester_key not in self.room_bookings:
            self.room_bookings[semester_key] = {}
        if slot_key not in self.room_bookings[semester_key]:
            self.room_bookings[semester_key][slot_key] = []
        self.room_bookings[semester_key][slot_key].append(booking)
    
    def _get_lab_rooms_assigned_to_other_section(self, semester_id, current_department):
        """Get set of lab room names already assigned to the other CSE section (CSE-A or CSE-B) this semester.
        Used so CSE-B gets different labs than CSE-A in the down table (each 80 students -> 2 labs each)."""
        other_section = 'CSE-A' if current_department == 'CSE-B' else ('CSE-B' if current_department == 'CSE-A' else None)
        if not other_section:
            return set()
        out = set()
        for key, room_list in self.assigned_lab_rooms.items():
            sem_id, dept, sess, code = key[0], key[1], key[2], key[3]
            if sem_id != semester_id or dept != other_section:
                continue
            for (_, _, room_name) in room_list:
                # room_name can be "106" or "106+107" (combined)
                for r in str(room_name).split('+'):
                    r = r.strip()
                    if r:
                        out.add(r)
        return out

    def _assign_room(self, day, slot, course_code, department, session, semester_id, is_lab=False, is_combined=False, required_capacity=0, slots=None, prefer_avoid_lab_rooms=None, ignore_conflicts=False):
        """Assign a room for a course at the specified slots with specific rules.
        prefer_avoid_lab_rooms: when assigning labs for CSE-B, set of lab room names already used by CSE-A (use last).
        ignore_conflicts: if True, skip availability checks and force a room allocation even if it conflicts.
        """
        semester_key = f"sem_{semester_id}"
        slot_sequence = slots if slots else [slot]
        prefer_avoid_lab_rooms = set(prefer_avoid_lab_rooms) if prefer_avoid_lab_rooms else set()
        
        def _room_available(room_name):
            """Check if a room is free globally (across ALL semesters) for all slots in slot_sequence,
            respecting session-based sharing (Pre-Mid and Post-Mid can share a room)."""
            if ignore_conflicts:
                return True
            for slot_label in slot_sequence:
                occ_key = (day, slot_label)
                if occ_key in self.room_occupancy:
                    # Current room usage at this slot: set of (room_name, session, course_code)
                    usage = self.room_occupancy[occ_key]
                    
                    target_sess = str(session).strip().lower()
                    target_code = str(course_code).strip().upper()
                    
                    for r_name, r_sess, r_code in usage:
                        if r_name == room_name:
                            # 1. Same course can always share the room
                            if str(r_code).strip().upper() == target_code:
                                continue
                            
                            existing_sess = str(r_sess).strip().lower()
                            # FULL session (or "full", "both") blocks everything else
                            if "full" in existing_sess or "both" in existing_sess:
                                return False
                            if "full" in target_sess or "both" in target_sess:
                                return False
                            
                            # Same session type blocks each other
                            if ("pre" in target_sess and "pre" in existing_sess):
                                return False
                            if ("post" in target_sess and "post" in existing_sess):
                                return False
            return True
        
        def _mark_room_usage(room_name, target_allocation):
            """Mark a room as used globally for all slots in slot_sequence with session info."""
            for slot_label in slot_sequence:
                occ_key = (day, slot_label)
                if occ_key not in self.room_occupancy:
                    self.room_occupancy[occ_key] = set()
                # Store tuple of (room_name, session, course_code) to allow sophisticated sharing logic
                self.room_occupancy[occ_key].add((room_name, session, course_code))
                # Log all rooms (including C004) so classroom allocation export can show usage
                if room_name:
                    self._log_room_booking(semester_id, day, slot_label, room_name, department, course_code, session)
                self.room_allocation_count[room_name] = self.room_allocation_count.get(room_name, 0) + 1
            target_allocation.append((day, slot, room_name))

        def _log_only_booking(room_name, target_allocation):
            """Log booking without adding to room_occupancy list (already marked by first dept in group)."""
            for slot_label in slot_sequence:
                if room_name:
                    self._log_room_booking(semester_id, day, slot_label, room_name, department, course_code, session)
            target_allocation.append((day, slot, room_name))
        
        # RULE 1: Combined classes MUST use C004
        if is_combined:
            if not self.c004_room:
                 print(f"      ERROR: C004 not found for combined class {course_code}")
                 return None
            room_name, room_capacity = self.c004_room
            
            # Check capacity
            if required_capacity > 0 and room_capacity < required_capacity:
                 print(f"      ERROR: C004 capacity ({room_capacity}) insufficient for combined class {course_code} ({required_capacity} students)")
                 return None
            
            allocation_key = (semester_id, department, session, course_code)
            if allocation_key not in self.assigned_rooms:
                self.assigned_rooms[allocation_key] = []
            
            # If C004 is NOT available (already booked)
            if not _room_available(room_name):
                # Is it booked by the SAME logical combined course?
                # MUST also be in the SAME combined group (e.g. CSE-A/B together, DSAI/ECE together)
                same_group_found = False
                for slot_label in slot_sequence:
                    slot_key = (day, slot_label)
                    # Use existing bookings to check who is in C004
                    if slot_key in self.room_bookings.get(semester_key, {}):
                        for b in self.room_bookings[semester_key][slot_key]:
                            if b.get('room') == room_name:
                                booked_course = str(b.get('course', '')).strip()
                                booked_dept = b.get('dept', '')
                                
                                # If it's the same course code AND the same combined group, it's valid sharing
                                if booked_course == str(course_code).strip() and \
                                   self._get_combined_group(booked_dept) == self._get_combined_group(department):
                                    same_group_found = True
                                    break
                    if same_group_found:
                        break
                
                if same_group_found:
                    _log_only_booking(room_name, self.assigned_rooms[allocation_key])
                    print(f"      Assigned C004 (shared) for combined class {course_code} at {day} {slot}")
                    return room_name

                else:
                    # Different course is in C004 -> Real conflict
                    print(f"      WARNING: C004 conflict at {day} {slot} - already used by another course; skipping {course_code}")
                    return None
            
            # Otherwise C004 is free
            _mark_room_usage(room_name, self.assigned_rooms[allocation_key])
            print(f"      Assigned C004 for combined class {course_code} at {day} {slot}")
            return room_name
        
        # RULE 2: Labs - Department specific assignment (single lab or consecutive labs when capacity exceeds one room)
        if is_lab:
            # CSE and DSAI get software labs
            if department in ['CSE-A', 'CSE-B', 'CSE', 'DSAI']:
                available_rooms = self.software_lab_rooms.copy()
                lab_type = "software"
            # ECE gets hardware labs
            elif department in ['ECE']:
                available_rooms = self.hardware_lab_rooms.copy()
                lab_type = "hardware"
            else:
                # Default to all labs if department doesn't match
                available_rooms = self.lab_rooms.copy()
                lab_type = "any"
            
            # Filter out rooms already occupied at this time
            filtered_rooms = []
            for name, cap in available_rooms:
                if _room_available(name):
                    filtered_rooms.append((name, cap))
            available_rooms = filtered_rooms
            
            allocation_key = (semester_id, department, session, course_code)
            if allocation_key not in self.assigned_lab_rooms:
                self.assigned_lab_rooms[allocation_key] = []
            
            # Try single lab first if one room fits
            if required_capacity > 0:
                fitting_single = [(name, cap) for name, cap in available_rooms if cap >= required_capacity]
                if fitting_single:
                    # Prefer least-used rooms (balance utilization), then avoid CSE-A labs for CSE-B, then capacity
                    fitting_single.sort(key=lambda x: (
                        self.room_allocation_count.get(x[0], 0),
                        x[0] in prefer_avoid_lab_rooms,
                        x[1], x[0]))
                    selected_room = fitting_single[0][0]
                    if not _room_available(selected_room):
                        for name, _ in fitting_single[1:]:
                            if _room_available(name):
                                selected_room = name
                                break
                    if not _room_available(selected_room):
                        print(f"      WARNING: No {lab_type} lab available for {course_code} at {day} {slot} (all occupied)")
                        return None
                    _mark_room_usage(selected_room, self.assigned_lab_rooms[allocation_key])
                    print(f"      Assigned {lab_type} lab {selected_room} for {course_code} at {day} {slot}")
                    return selected_room
                # No single lab fits: assign consecutive labs (e.g. 106+107) so combined capacity >= required_capacity
                # Sort: prefer least-used, then avoid CSE-A labs for CSE-B, then by name for consecutive numbers
                def _lab_sort_key(item):
                    name, cap = item
                    n = name.upper() if isinstance(name, str) else str(name)
                    return (self.room_allocation_count.get(name, 0), name in prefer_avoid_lab_rooms, n, cap)
                available_rooms.sort(key=_lab_sort_key)
                total_cap = 0
                selected_rooms = []
                for name, cap in available_rooms:
                    if total_cap >= required_capacity:
                        break
                    if _room_available(name):
                        selected_rooms.append((name, cap))
                        total_cap += cap
                if selected_rooms and total_cap >= required_capacity:
                    for name, cap in selected_rooms:
                        if not _room_available(name):
                            print(f"      WARNING: Lab {name} no longer available for {course_code} at {day} {slot}")
                            return None
                    for name, cap in selected_rooms:
                        _mark_room_usage(name, self.assigned_lab_rooms[allocation_key])
                    combined = '+'.join(r[0] for r in selected_rooms)
                    print(f"      Assigned {lab_type} labs {combined} (capacity {total_cap} for {required_capacity} students) for {course_code} at {day} {slot}")
                    return combined
            else:
                # No capacity requirement: prefer least-used room, then avoid prefer_avoid_lab_rooms
                if available_rooms:
                    available_rooms.sort(key=lambda x: (
                        self.room_allocation_count.get(x[0], 0), x[0] in prefer_avoid_lab_rooms, x[1], x[0]))
                    selected_room = available_rooms[0][0]
                    if not _room_available(selected_room):
                        for name, _ in available_rooms[1:]:
                            if _room_available(name):
                                selected_room = name
                                break
                    if not _room_available(selected_room):
                        print(f"      WARNING: No {lab_type} lab available for {course_code} at {day} {slot}")
                        return None
                    _mark_room_usage(selected_room, self.assigned_lab_rooms[allocation_key])
                    print(f"      Assigned {lab_type} lab {selected_room} for {course_code} at {day} {slot}")
                    return selected_room
            print(f"      WARNING: No {lab_type} lab(s) available for {course_code} at {day} {slot} (need capacity {required_capacity})")
            return None
        
        # RULE 3: Regular classes - Use normal classrooms (NOT C004).
        # IMPORTANT: Follow the order from the input Excel sheet.
        available_rooms = [(name, cap) for name, cap in self.nonlab_rooms if name.upper() != 'C004']
        # Keep only rooms that are currently free for the whole slot sequence
        available_rooms = [(name, cap) for name, cap in available_rooms if _room_available(name)]

        if required_capacity > 0:
            # Keep order, only filter by capacity
            available_rooms = [(name, cap) for name, cap in available_rooms if cap >= required_capacity]

        if available_rooms:
            # Pick the first suitable room according to input order
            selected_room = None
            for name, _ in available_rooms:
                if _room_available(name):
                    selected_room = name
                    break

            if not selected_room:
                print(f"      WARNING: No classroom available for {course_code} at {day} {slot} (all occupied)")
                return None

            allocation_key = (semester_id, department, session, course_code)
            if allocation_key not in self.assigned_rooms:
                self.assigned_rooms[allocation_key] = []
            _mark_room_usage(selected_room, self.assigned_rooms[allocation_key])
            return selected_room
        return None
    
    def _schedule_minor_classes(self, schedule, department, session, semester_id):
        """Schedule Minor subject classes ONLY in configured MINOR_SLOTS (e.g., 07:30-08:30 split).
        All departments/sections in a semester get the same minor slots."""
        # Skip minor scheduling entirely for semester 1 as per requirement
        if int(semester_id) == 1:
            return
        scheduled = 0
        attempts = 0
        max_attempts = 200
        
        # Compute valid minor start slots (so MINOR_DURATION consecutive slots are within MINOR_SLOTS)
        minor_starts = []
        for s in MINOR_SLOTS:
            seq = self._get_consecutive_slots(s, MINOR_DURATION)
            if len(seq) == MINOR_DURATION and all(x in MINOR_SLOTS for x in seq):
                minor_starts.append(s)
        
        if not minor_starts:
            # nothing to schedule if config mismatched
            return
        
        semester_key = f"sem_{semester_id}"
        # If already assigned for this semester, use the same slots
        if semester_key in self.semester_minor_slots:
            assigned = self.semester_minor_slots[semester_key]
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, MINOR_DURATION)
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
            return

        assigned = []
        while scheduled < MINOR_CLASSES_PER_WEEK and attempts < max_attempts:
            attempts += 1
            
            day = random.choice(DAYS)
            start = random.choice(minor_starts)
            slots = self._get_consecutive_slots(start, MINOR_DURATION)
            
            # IMPORTANT: Check if any slot is reserved by higher priority courses (Electives, Combined)
            slot_reserved = False
            for slot in slots:
                if self._is_slot_reserved_global(day, slot, semester_id, 'Minor'):
                    slot_reserved = True
                    break
            
            if slot_reserved:
                # Skip this slot as it conflicts with higher priority courses
                continue
            
            if (len(slots) == MINOR_DURATION and
                self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                
                assigned.append((day, start))
                scheduled += 1
        # Save assigned slots for all departments in this semester
        if assigned:
            self.semester_minor_slots[semester_key] = assigned

    def _get_combined_group(self, dept_label, course_code=None):
        """Determine which departments can share combined slots."""
        # Universal courses (HSS or in FORCED_COMBINED) always group ALL departments
        if course_code:
            code_upper = str(course_code).strip().upper()
            if 'HSS' in code_upper or code_upper in FORCED_COMBINED_COURSES:
                return 'ALL'
        
        # Standard combined grouping
        if dept_label in {'CSE-A', 'CSE-B'}:
            return 'CSE'
        if dept_label in {'DSAI', 'ECE'}:
            return 'DSAI_ECE'
        return None

    def _find_combined_slots(self, schedule, course_code, component, duration, department, session, semester_id, avoid_days=None):
        """Find available slots for combined classes across all departments in the same group.
        Returns list of (day, start_slot) tuples for combined scheduling."""
        combined_slots = []
        group_key = self._get_combined_group(department)
        target_code = str(course_code).strip().upper()
        
        all_possible_starts = [s for s in TEACHING_SLOTS if s not in LUNCH_SLOTS]
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        max_attempts = 1500
        attempts = 0
        avoid_days = set(avoid_days) if avoid_days else set()
        
        while attempts < max_attempts:
            attempts += 1
            day = random.choice([d for d in DAYS if d not in avoid_days])
            if not all_possible_starts:
                break
            start_slot = random.choice(all_possible_starts)
            slots = self._get_consecutive_slots(start_slot, duration)
            slots = [slot for slot in slots if slot in regular_slots]
            
            if len(slots) == duration:
                # IMPORTANT: For combined classes, check if slots are reserved by Electives (Priority 1)
                slot_reserved = False
                for slot in slots:
                    if self._is_slot_reserved_global(day, slot, semester_id, 'Combined'):
                        slot_reserved = True
                        break
                
                if slot_reserved:
                    continue

                # Check if slots are available for ALL departments in the group
                all_available = True
                for dept in DEPARTMENTS:
                    if self._get_combined_group(dept, course_code) == group_key:
                        # Check local availability for each department
                        # Note: 'schedule' arg is only for current department, but we can't easily access other dept schedules here.
                        # We rely on global slots mostly.
                        
                        # Check global availability
                        if not self._is_time_slot_available_global(day, slots, dept, session, semester_id):
                            all_available = False
                            break
                            
                if all_available:
                     # Additional check: Is C004 free for these slots?
                     # Since this is for combined classes, we KNOW we need C004.
                     c004_free = True
                     if self.c004_room:
                         c004_name = self.c004_room[0]
                         for s in slots:
                             occ_key = (day, s)
                             if occ_key in self.room_occupancy:
                                 target_sess = str(session).strip().lower()
                                 for r_name, r_sess, r_code in self.room_occupancy[occ_key]:
                                     if r_name == c004_name:
                                         # 1. Same course can share
                                         if str(r_code).strip().upper() == target_code:
                                             continue
                                         
                                         # 2. Different sessions can share
                                         existing_sess = str(r_sess).strip().lower()
                                         if ("full" in existing_sess or "both" in existing_sess or
                                             "full" in target_sess or "both" in target_sess or
                                             ("pre" in target_sess and "pre" in existing_sess) or
                                             ("post" in target_sess and "post" in existing_sess)):
                                             c004_free = False
                                             break
                                 if not c004_free:
                                     break
                     
                     if c004_free:
                        combined_slots.append((day, start_slot))
                        break
        
        return combined_slots

    def _schedule_combined_class(self, schedule, course_code, component, duration, required_count, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule combined class for all departments in the same group with C004 room allocation."""
        scheduled_count = 0
        scheduled_slots = []
        assigned_rooms = []
        
        # Check if combined slots already exist for this course
        group_key = self._get_combined_group(department, course_code)
        course_key = str(course_code).strip()
        is_lab = (component == 'Lab')
        
        # Check global combined slots first
        if group_key:
            global_key = ('GLOBAL', group_key, course_key, component)
            assigned = self.global_combined_course_slots.get(global_key, [])
            
            if assigned and len(assigned) >= required_count:
                # Use existing global slots
                for day, start in assigned[:required_count]:
                    slots = self._get_consecutive_slots(start, duration)
                    self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                    self._mark_slots_busy_global(day, slots, department, session, semester_id)
                    scheduled_slots.extend([(day, slot) for slot in slots])
                    
                    # Assign C004 for this combined class
                    room = self._assign_room(day, start, course_code, department, session, semester_id, 
                                          is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
                    if room:
                        assigned_rooms.append(room)
                    
                    scheduled_count += 1
                return scheduled_count, scheduled_slots, assigned_rooms
            
            # Check semester-specific combined slots (restricted by group)
            sem_key = (semester_id, group_key, course_key, component)
            assigned = self.semester_combined_course_slots.get(sem_key, [])
            
            if assigned and len(assigned) >= required_count:
                # Use existing semester slots
                for day, start in assigned[:required_count]:
                    slots = self._get_consecutive_slots(start, duration)
                    self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                    self._mark_slots_busy_global(day, slots, department, session, semester_id)
                    scheduled_slots.extend([(day, slot) for slot in slots])
                    
                    # Assign C004 for this combined class (shared within group)
                    room = self._assign_room(day, start, course_code, department, session, semester_id, 
                                          is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
                    if room:
                        assigned_rooms.append(room)
                    
                    scheduled_count += 1
                return scheduled_count, scheduled_slots, assigned_rooms
        
        # Find new combined slots
        while scheduled_count < required_count:
            combined_slots = self._find_combined_slots(schedule, course_code, component, duration, department, session, semester_id, avoid_days)
            
            if not combined_slots:
                break
            
            day, start_slot = combined_slots[0]
            slots = self._get_consecutive_slots(start_slot, duration)
            
            # Assign C004 first (before marking slots globally)
            room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                  is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
            
            if room:  # C004 must be available for combined classes
                # Mark slots for all departments in group
                for dept in DEPARTMENTS:
                    if self._get_combined_group(dept, course_code) == group_key:
                        self._mark_slots_busy_global(day, slots, dept, session, semester_id)
                
                # Mark locally for current department
                self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                scheduled_slots.extend([(day, slot) for slot in slots])
                scheduled_count += 1
                assigned_rooms.append(room)
                
                # Store combined slot
                if group_key:
                    global_key = ('GLOBAL', group_key, course_key, component)
                    if global_key not in self.global_combined_course_slots:
                        self.global_combined_course_slots[global_key] = []
                    self.global_combined_course_slots[global_key].append((day, start_slot))
                    
                    sem_key = (semester_id, group_key, course_key, component)
                    if sem_key not in self.semester_combined_course_slots:
                        self.semester_combined_course_slots[sem_key] = []
                    self.semester_combined_course_slots[sem_key].append((day, start_slot))
            else:
                # C004 not available, try different slot
                continue
        
        return scheduled_count, scheduled_slots, assigned_rooms

    def _schedule_lectures(self, schedule, course_code, lectures_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0):
        """Schedule lecture sessions with room allocation, returns list of (day, slot) tuples.
        Prioritizes slots ending at :30, falls back to remaining slots if needed."""
        if lectures_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 20000  # Higher attempts for better allocation percentage
        used_days = set()
        if avoid_days is None:
            avoid_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Lecture', LECTURE_DURATION, 
                lectures_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= lectures_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining
        
        # PRIORITY 2: Handle electives with common slots
        if is_elective:
            elective_slots = self._schedule_elective_classes(schedule, course_code, lectures_per_week, department, session, semester_id, avoid_days, required_capacity)
            if elective_slots:
                return elective_slots
        
        skip_room_assignment = is_elective or is_minor

        # PRIORITY 3: Regular lecture scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LECTURE_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        while len(scheduled_slots) < lectures_per_week * LECTURE_DURATION and attempts < max_attempts:
            attempts += 1
            
            # Try to use days where this course isn't already scheduled (strict same-day rule)
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                # If no perfect combo, try to at least avoid multiple lectures on same day (lec+lec rule)
                available_combos = [combo for combo in all_combinations if combo[0] not in used_days]
            
            if not available_combos:
                # Last resort: allow any day if absolutely needed to avoid failure
                available_combos = all_combinations
            
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            # IMPORTANT: For non-elective courses, check if any slot is reserved by higher priority courses
            priority_type = 'Lecture'
            if is_elective:
                priority_type = 'Elective'
            elif is_combined:
                priority_type = 'Combined'
            elif is_minor:
                priority_type = 'Minor'
                
            if priority_type not in ['Elective', 'Combined']:
                slot_reserved = False
                for slot in slots:
                    if self._is_slot_reserved_global(day, slot, semester_id, priority_type):
                        slot_reserved = True
                        break
                
                if slot_reserved:
                    # Skip this combination as it conflicts with higher priority slots
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
                    continue
            
            # Check all slots are available (both local and global)
            slots_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    slots_available = False
                    break
            
            if slots_available:
                room_available = True
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                             is_lab=False, is_combined=False, required_capacity=required_capacity, slots=slots)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < lectures_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{lectures_per_week} lectures (attempts: {attempts})")
        return scheduled_slots

    def _schedule_elective_classes(self, schedule, course_code, elective_per_week, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule elective classes at the same slot/day for all departments/sections in a semester.
        ALL electives in a semester must use the same time slots across ALL departments (CSE, DSAI, ECE).
        Returns list of (day, slot) tuples."""
        if elective_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 20000  # Higher for better allocation percentage (increased from 2k)
        semester_key = f"sem_{semester_id}"
        avoid_days = set(avoid_days or [])
        
        # Determine elective group for more granular slotting (e.g. 5-ELEC1 -> ELEC1)
        elective_group = 'ALL_ELECTIVES'
        if 'ELEC' in str(course_code).upper():
             import re
             m = re.search(r'ELEC\d+', str(course_code).upper())
             if m:
                 elective_group = m.group(0)

        # Use a common key for specific elective group in a semester to ensure same time slots
        common_elective_key = (semester_key, elective_group)
        elective_key = (semester_key, course_code)
        used_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # First, check if common elective slots have been assigned for this semester
        # If yes, ALL electives must use those same slots
        if common_elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[common_elective_key]
            print(f"      Using common elective slots for {course_code} (already assigned for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            return scheduled_slots
        
        # If no common slots assigned yet, check if this specific course has slots assigned
        # (This handles legacy cases where course-specific slots were assigned first)
        if elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[elective_key]
            # Promote to common slots so all future electives use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            print(f"      Using existing elective slots for {course_code} (promoted to common slots for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            return scheduled_slots

        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LECTURE_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        assigned = []
        scheduled = 0
        while scheduled < elective_per_week and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos) if available_combos else (None, None, None)
            if day is None:
                break
            
            # Check if this slot is available for the current department/session
            if not (self._is_time_slot_available_local(schedule, day, slots) and
                    self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                continue
            
            # IMPORTANT: For electives, also check if this slot conflicts with NML courses
            # across ALL departments in this semester (both Pre-Mid and Post-Mid sessions)
            # This ensures electives don't overlap with NML courses
            conflicts_with_nml = False
            for dept in DEPARTMENTS:
                for sess in [PRE_MID, POST_MID]:
                    # Check if this slot is occupied by NML courses for any department/session
                    if not self._is_time_slot_available_global(day, slots, dept, sess, semester_id):
                        # Check if the conflict is with a Minor course
                        dept_key = f"{dept}_{sess}"
                        sem_key = f"sem_{semester_id}"
                        if sem_key in self.semester_global_slots and dept_key in self.semester_global_slots[sem_key]:
                            for slot in slots:
                                if (day, slot) in self.semester_global_slots[sem_key][dept_key]:
                                    # This slot is occupied - check if it's a Minor course
                                    # Minor courses are in MINOR_SLOTS
                                    if slot in MINOR_SLOTS:
                                        conflicts_with_nml = True
                                        break
                        if conflicts_with_nml:
                            break
                if conflicts_with_nml:
                    break
            
            if conflicts_with_nml:
                # Skip this slot as it conflicts with NML courses
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
                continue
            
            # Slot is available and doesn't conflict with NML courses
            if True:
                
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                assigned.append((day, start_slot))
                used_days.add(day)
                avoid_days.add(day)
                scheduled += 1
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        if assigned:
            # Store under common key so ALL electives in this semester use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            print(f"      Assigned common elective slots for semester {semester_id}: {assigned}")
        
        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < elective_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{elective_per_week} elective classes (attempts: {attempts})")
        return scheduled_slots

    def _schedule_elective_tutorials(self, schedule, course_code, elective_tutorials_per_week, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule elective tutorials at the same slot/day for all departments/sections in a semester.
        ALL elective tutorials in a semester must use the same time slots across ALL departments (CSE-A, CSE-B, DSAI, ECE)
        for both Pre-Mid and Post-Mid sessions.
        Returns list of (day, slot) tuples."""
        if elective_tutorials_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 20000  # Higher for better allocation percentage (increased from 2k)
        semester_key = f"sem_{semester_id}"
        avoid_days = set(avoid_days or [])
        
        # Determine elective group for more granular slotting (e.g. 5-ELEC1 -> ELEC1)
        elective_group = 'ALL_ELECTIVE_TUTORIALS'
        if 'ELEC' in str(course_code).upper():
             import re
             m = re.search(r'ELEC\d+', str(course_code).upper())
             if m:
                 elective_group = f"{m.group(0)}_TUTORIALS"

        # Use a common key for ALL elective tutorials of the same group in a semester 
        common_elective_tutorial_key = (semester_key, elective_group)
        elective_tutorial_key = (semester_key, course_code, 'Tutorial')
        used_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # First, check if common elective tutorial slots have been assigned for this semester
        # If yes, ALL elective tutorials must use those same slots (for both Pre-Mid and Post-Mid)
        if common_elective_tutorial_key in self.semester_elective_tutorial_slots:
            assigned = self.semester_elective_tutorial_slots[common_elective_tutorial_key]
            print(f"      Using common elective tutorial slots for {course_code} (already assigned for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, TUTORIAL_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            # Also store under course-specific key for backward compatibility
            self.semester_elective_tutorial_slots[elective_tutorial_key] = assigned
            return scheduled_slots
        
        # If no common slots assigned yet, check if this specific course has tutorial slots assigned
        # (This handles legacy cases where course-specific slots were assigned first)
        if elective_tutorial_key in self.semester_elective_tutorial_slots:
            assigned = self.semester_elective_tutorial_slots[elective_tutorial_key]
            # Promote to common slots so all future elective tutorials use the same slots
            self.semester_elective_tutorial_slots[common_elective_tutorial_key] = assigned
            print(f"      Using existing elective tutorial slots for {course_code} (promoted to common slots for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, TUTORIAL_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            return scheduled_slots

        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(TUTORIAL_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        assigned = []
        scheduled = 0
        while scheduled < elective_tutorials_per_week and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos) if available_combos else (None, None, None)
            if day is None:
                break
            
            # Check if this slot is available for the current department/session
            if not (self._is_time_slot_available_local(schedule, day, slots) and
                    self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                continue
            
            # IMPORTANT: For elective tutorials, also check if this slot conflicts with NML courses
            # across ALL departments in this semester (both Pre-Mid and Post-Mid sessions)
            # This ensures elective tutorials don't overlap with NML courses
            conflicts_with_nml = False
            for dept in DEPARTMENTS:
                for sess in [PRE_MID, POST_MID]:
                    # Check if this slot is occupied by NML courses for any department/session
                    if not self._is_time_slot_available_global(day, slots, dept, sess, semester_id):
                        # Check if the conflict is with a Minor course
                        dept_key = f"{dept}_{sess}"
                        sem_key = f"sem_{semester_id}"
                        if sem_key in self.semester_global_slots and dept_key in self.semester_global_slots[sem_key]:
                            for slot in slots:
                                if (day, slot) in self.semester_global_slots[sem_key][dept_key]:
                                    # This slot is occupied - check if it's a Minor course
                                    # Minor courses are in MINOR_SLOTS
                                    if slot in MINOR_SLOTS:
                                        conflicts_with_nml = True
                                        break
                        if conflicts_with_nml:
                            break
                if conflicts_with_nml:
                    break
            
            if conflicts_with_nml:
                # Skip this slot as it conflicts with NML courses
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
                continue
            
            # Slot is available and doesn't conflict with NML courses
            if True:
                
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                assigned.append((day, start_slot))
                used_days.add(day)
                avoid_days.add(day)
                scheduled += 1
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        if assigned:
            # Store under common key so ALL elective tutorials in this semester use the same slots
            # This works across all departments and both Pre-Mid and Post-Mid sessions
            self.semester_elective_tutorial_slots[common_elective_tutorial_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_tutorial_slots[elective_tutorial_key] = assigned
            print(f"      Assigned common elective tutorial slots for semester {semester_id}: {assigned}")
        
        scheduled_count = len(scheduled_slots) // TUTORIAL_DURATION
        if scheduled_count < elective_tutorials_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{elective_tutorials_per_week} elective tutorials (attempts: {attempts})")
        return scheduled_slots

    def _schedule_tutorials(self, schedule, course_code, tutorials_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0):
        """Schedule tutorial sessions with room allocation, returns list of (day, slot) tuples.
        Prioritizes slots ending at :30, falls back to remaining slots if needed."""
        if tutorials_per_week == 0:
            return []

        scheduled_slots = []
        attempts = 0
        max_attempts = 20000  # Increased for allocation success
        used_days = set()
        if avoid_days is None:
            avoid_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Tutorial', TUTORIAL_DURATION, 
                tutorials_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= tutorials_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining
        
        skip_room_assignment = is_elective or is_minor

        # PRIORITY 2: Handle electives with common slots
        if is_elective:
            elective_slots = self._schedule_elective_tutorials(schedule, course_code, tutorials_per_week, department, session, semester_id, avoid_days, required_capacity)
            if elective_slots:
                return elective_slots

        # PRIORITY 3: Regular tutorial scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(TUTORIAL_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        while len(scheduled_slots) < tutorials_per_week * TUTORIAL_DURATION and attempts < max_attempts:
            attempts += 1
            # Try to use days where this course isn't already scheduled (strict same-day rule: avoid Lec and Tut)
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                # If no perfect combo, try to at least avoid multiple tutorials on same day (tut+tut rule)
                available_combos = [combo for combo in all_combinations if combo[0] not in used_days]
            
            if not available_combos:
                available_combos = all_combinations
            
            if not available_combos:
                break

            day, start_slot, slots = random.choice(available_combos)
            
            # IMPORTANT: For non-elective courses, check if any slot is reserved by higher priority courses
            priority_type = 'Tutorial'
            if is_elective:
                priority_type = 'Elective'
            elif is_combined:
                priority_type = 'Combined'
            elif is_minor:
                priority_type = 'Minor'

            if priority_type not in ['Elective', 'Combined']:
                slot_reserved = False
                for slot in slots:
                    if self._is_slot_reserved_global(day, slot, semester_id, priority_type):
                        slot_reserved = True
                        break
                
                if slot_reserved:
                    # Skip this combination as it conflicts with higher priority slots
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
                    continue
            
            room_available = True
            slots_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                        self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    slots_available = False
                    break
            
            if slots_available:
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                              is_lab=False, is_combined=False, required_capacity=required_capacity, slots=slots)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // TUTORIAL_DURATION
        if scheduled_count < tutorials_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{tutorials_per_week} tutorials (attempts: {attempts})")
        
        return scheduled_slots

    def _schedule_labs(self, schedule, course_code, labs_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0, prefer_avoid_lab_rooms=None):
        """Schedule lab sessions in regular time slots (multi-slot labs) with department-specific lab allocation.
        Prioritizes slots ending at :30, falls back to remaining slots if needed.
        Returns list of (day, slot) tuples."""
        if labs_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 20000  # Increased for lab allocation robustness
        used_days = set()
        if avoid_days is None:
            avoid_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Lab', LAB_DURATION, 
                labs_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= labs_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining

        skip_room_assignment = is_elective or is_minor

        # PRIORITY 2: Department-specific lab scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LAB_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LAB_DURATION)
                if len(slots) == LAB_DURATION and all(s in regular_slots for s in slots):
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LAB_DURATION)
                if len(slots) == LAB_DURATION and all(s in regular_slots for s in slots):
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations
        
        while len(scheduled_slots) < labs_per_week * LAB_DURATION and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            # IMPORTANT: For non-elective/non-combined courses, check if any slot is reserved by higher priority courses
            priority_type = 'Lab'
            if is_elective:
                priority_type = 'Elective'
            elif is_combined:
                priority_type = 'Combined'
            elif is_minor:
                priority_type = 'Minor'

            if priority_type not in ['Elective', 'Combined']:
                slot_reserved = False
                for slot in slots:
                    if self._is_slot_reserved_global(day, slot, semester_id, priority_type):
                        slot_reserved = True
                        break
                
                if slot_reserved:
                    # Skip this combination as it conflicts with higher priority slots
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
                    continue
            
            room_available = True
            # Check if all slots are available
            all_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    all_available = False
                    break
            
            if all_available:
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                             is_lab=True, is_combined=False, required_capacity=required_capacity, slots=slots, prefer_avoid_lab_rooms=prefer_avoid_lab_rooms)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lab')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // LAB_DURATION
        if scheduled_count < labs_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{labs_per_week} labs (attempts: {attempts})")
        return scheduled_slots

    def _force_allocate(self, schedule, course_code, count, duration, department, session, semester_id, component, required_capacity=0):
        """Forcefully allocate remaining sessions for a course component.
        This ignores global slot reservations and simply fills any locally free slots.
        It will still mark slots busy globally and attempt to allocate a room (with
        conflicts permitted).
        Returns list of (day, slot) tuples assigned."""
        slots_assigned = []
        if count <= 0:
            return slots_assigned
        # first pass: try available local slots
        for day in DAYS:
            for start in TEACHING_SLOTS:
                if len(slots_assigned) >= count * duration:
                    break
                seq = self._get_consecutive_slots(start, duration)
                if len(seq) != duration:
                    continue
                # only require local availability
                if self._is_time_slot_available_local(schedule, day, seq):
                    # assign room ignoring conflicts
                    room = self._assign_room(day, start, course_code, department, session, semester_id,
                                              is_lab=(component == 'Lab'), is_combined=False,
                                              required_capacity=required_capacity, slots=seq,
                                              ignore_conflicts=True)
                    for slot in seq:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, component)
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        slots_assigned.append((day, slot))
            if len(slots_assigned) >= count * duration:
                break
        # second pass: if still not enough, overwrite any slots
        if len(slots_assigned) < count * duration:
            for day in DAYS:
                for start in TEACHING_SLOTS:
                    if len(slots_assigned) >= count * duration:
                        break
                    seq = self._get_consecutive_slots(start, duration)
                    if len(seq) != duration:
                        continue
                    # force even if not free
                    room = self._assign_room(day, start, course_code, department, session, semester_id,
                                              is_lab=(component == 'Lab'), is_combined=False,
                                              required_capacity=required_capacity, slots=seq,
                                              ignore_conflicts=True)
                    for slot in seq:
                        # mark regardless of previous content
                        schedule.loc[day, slot] = f"{course_code}{' (Lab)' if component=='Lab' else ''}"
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        slots_assigned.append((day, slot))
                if len(slots_assigned) >= count * duration:
                    break
        if slots_assigned:
            print(f"      INFO: Forced allocation of {len(slots_assigned)//duration} {component.lower()}(s) for {course_code}")
        return slots_assigned

    def _schedule_course(self, schedule, course, department, session, semester_id, prefer_avoid_lab_rooms=None):
        """Schedule all components of a course based on LTPSC with proper room allocation."""
        course_code = course['Course Code']
        lectures_per_week = course['Lectures_Per_Week']
        tutorials_per_week = course['Tutorials_Per_Week']
        labs_per_week = course['Labs_Per_Week']
        
        # Robust elective detection: check multiple column name variants + pattern overrides
        elective_flag = False
        for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
            if colname in course.index:
                elective_flag = str(course.get(colname, '')).upper() == 'YES'
                if elective_flag:
                    break
        # Pattern overrides: force ELEC as elective; force HSS as not elective
        course_code_str = str(course.get('Course Code', '')).upper()
        course_name_str = str(course.get('Course Name', '')).upper()
        if 'ELEC' in course_code_str or 'ELEC' in course_name_str:
            elective_flag = True
        if 'HSS' in course_code_str or 'HSS' in course_name_str:
            elective_flag = False
        # Determine if HSS
        is_hss = ('HSS' in course_code_str) or ('HSS' in course_name_str)
        is_minor_course = (MINOR_SUBJECT.upper() in course_code_str) or (MINOR_SUBJECT.upper() in course_name_str)
        
        # Read Combined Class column (handle multiple column name variants)
        combined_class_flag = False
        for colname in ['Combined Class', 'COMBINED CLASS', 'Combined Class ', 'COMBINED CLASS ']:
            if colname in course.index:
                combined_class_val = str(course.get(colname, '')).strip().upper()
                combined_class_flag = combined_class_val == 'YES'
                if combined_class_flag:
                    break
        # Forced combined override from config (e.g., EC161)
        if course_code_str in FORCED_COMBINED_COURSES:
            combined_class_flag = True
        
        # Get student count for room capacity
        registered_students = 0
        if 'Registered Students' in course.index:
            try:
                registered_students = int(float(course.get('Registered Students', 0)))
            except:
                registered_students = 0
        # CSE total 160 = CSE-A 80 + CSE-B 80: each section gets only 2 labs (capacity 80), not 4
        if department in ('CSE-A', 'CSE-B') and registered_students > CSE_SECTION_CAPACITY:
            registered_students = CSE_SECTION_CAPACITY
        
        # IMPORTANT: Follow LTPSC strictly for ALL courses (including electives)
        # Do not override parsed weekly counts; use values from Excel (via parse_ltpsc)
        elective_status = " [ELECTIVE]" if elective_flag else ""
        combined_status = " [COMBINED]" if combined_class_flag else ""
        print(f"      Scheduling {course_code}{elective_status}{combined_status}: L={lectures_per_week}, T={tutorials_per_week}, P={labs_per_week}, Students={registered_students}")
        
        # Track used days for this course scoped to semester+department+session
        scoped_key = (semester_id, department, session, course_code)
        if scoped_key not in self.scheduled_courses:
            self.scheduled_courses[scoped_key] = set()

        success_counts = {'lectures': 0, 'tutorials': 0, 'labs': 0}
        scheduled_slots = []
        assigned_rooms = []
        assigned_lab_rooms = []

        # PRIORITY 1: Schedule combined classes first
        # This ensures combined classes get priority for available slots
        avoid_days = set()
        
        # PRIORITY 1: Schedule Labs
        if labs_per_week > 0:
            lab_avoid_days = set() 
            lab_slots = self._schedule_labs(
                schedule, course_code, labs_per_week, department, session, semester_id,
                lab_avoid_days, is_combined=combined_class_flag,
                is_elective=elective_flag, is_minor=is_minor_course,
                required_capacity=registered_students,
                prefer_avoid_lab_rooms=prefer_avoid_lab_rooms
            )
            scheduled_slots.extend(lab_slots)
            success_counts['labs'] = len(lab_slots) // LAB_DURATION
            # fallback if not all labs scheduled
            if success_counts['labs'] < labs_per_week:
                missing = labs_per_week - success_counts['labs']
                forced = self._force_allocate(schedule, course_code, missing, LAB_DURATION,
                                              department, session, semester_id, 'Lab', registered_students)
                scheduled_slots.extend(forced)
                success_counts['labs'] += len(forced) // LAB_DURATION

        # PRIORITY 2: Schedule Lectures
        if lectures_per_week > 0:
            lecture_slots = self._schedule_lectures(
                schedule, course_code, lectures_per_week, department, session, semester_id,
                avoid_days, is_combined=combined_class_flag, is_elective=elective_flag, 
                is_minor=is_minor_course, required_capacity=registered_students
            )
            scheduled_slots.extend(lecture_slots)
            success_counts['lectures'] = len(lecture_slots) // LECTURE_DURATION
            # fallback if not all lectures scheduled
            if success_counts['lectures'] < lectures_per_week:
                missing = lectures_per_week - success_counts['lectures']
                forced = self._force_allocate(schedule, course_code, missing, LECTURE_DURATION,
                                              department, session, semester_id, 'Lecture', registered_students)
                scheduled_slots.extend(forced)
                success_counts['lectures'] += len(forced) // LECTURE_DURATION
        
        # PRIORITY 3: Schedule Tutorials
        if tutorials_per_week > 0:
            tutorial_slots = self._schedule_tutorials(
                schedule, course_code, tutorials_per_week, department, session, semester_id,
                avoid_days, is_combined=combined_class_flag, is_elective=elective_flag,
                is_minor=is_minor_course, required_capacity=registered_students
            )
            scheduled_slots.extend(tutorial_slots)
            success_counts['tutorials'] = len(tutorial_slots) // TUTORIAL_DURATION
            # fallback if not all tutorials scheduled
            if success_counts['tutorials'] < tutorials_per_week:
                missing = tutorials_per_week - success_counts['tutorials']
                forced = self._force_allocate(schedule, course_code, missing, TUTORIAL_DURATION,
                                              department, session, semester_id, 'Tutorial', registered_students)
                scheduled_slots.extend(forced)
                success_counts['tutorials'] += len(forced) // TUTORIAL_DURATION

        # Get assigned rooms for this course
        allocation_key = (semester_id, department, session, course_code)
        if allocation_key in self.assigned_rooms:
            assigned_rooms = [room_info[2] for room_info in self.assigned_rooms[allocation_key]]
        if allocation_key in self.assigned_lab_rooms:
            lab_room_list = [room_info[2] for room_info in self.assigned_lab_rooms[allocation_key]]
            # Show combined labs as "106+107" when multiple labs assigned for capacity (order preserved, unique)
            lab_room_display = '+'.join(dict.fromkeys(lab_room_list)) if lab_room_list else ''
        else:
            lab_room_display = ''

        # Store actual allocation counts with room information
        self.actual_allocations[allocation_key] = {
            'lectures': success_counts['lectures'],
            'tutorials': success_counts['tutorials'],
            'labs': success_counts['labs'],
            'combined_class': combined_class_flag,
            'room': assigned_rooms[0] if assigned_rooms else '',
            'lab_room': lab_room_display
        }

        # Store scheduled slots for conflict tracking
        if scoped_key not in self.scheduled_slots:
            self.scheduled_slots[scoped_key] = []
        self.scheduled_slots[scoped_key].extend(scheduled_slots)
        
        # Track days used by this course
        for day, slot in scheduled_slots:
            self.scheduled_courses[scoped_key].add(day)

        return success_counts

    def generate_department_schedule(self, semester_id, department, session):
        """Generate a complete weekly schedule for a department and session."""
        print(f"\nGenerating schedule for {department} {session} (Semester {semester_id})")
        
        # Initialize empty schedule
        schedule = self._initialize_schedule()
        
        # Get courses for this department and session
        sem_courses = ExcelLoader.get_semester_courses(self.dfs, semester_id)
        if sem_courses.empty:
            print(f"WARNING: No courses found for semester {semester_id}")
            return schedule
        
        # Parse LTPSC
        sem_courses = ExcelLoader.parse_ltpsc(sem_courses)
        if sem_courses.empty:
            print(f"WARNING: No valid courses after LTPSC parsing for semester {semester_id}")
            return schedule
        
        # Filter for department
        if 'Department' in sem_courses.columns:
            # Include courses matching exact department OR 'ALL'
            dept_mask = sem_courses['Department'].astype(str).str.contains(f"^{department}$|^ALL$", na=False, regex=True)
            dept_courses = sem_courses[dept_mask].copy()
        else:
            dept_courses = sem_courses.copy()
        
        if dept_courses.empty:
            print(f"WARNING: No courses found for {department} in semester {semester_id}")
            return schedule
        
        # Divide by session
        pre_mid_courses, post_mid_courses = ExcelLoader.divide_courses_by_session(dept_courses, department, all_sem_courses=sem_courses)
        
        # Select appropriate session
        if session == PRE_MID:
            session_courses = pre_mid_courses
        else:
            session_courses = post_mid_courses
        
        if session_courses.empty:
            print(f"WARNING: No courses assigned to {department} {session} session")
            return schedule
        
        # For CSE-B (and CSE-A if scheduled after CSE-B), prefer labs not used by the other section so down table shows different labs
        prefer_avoid_lab_rooms = set()
        if department in ('CSE-A', 'CSE-B'):
            prefer_avoid_lab_rooms = self._get_lab_rooms_assigned_to_other_section(semester_id, department)
            if prefer_avoid_lab_rooms:
                print(f"  Prefer labs different from other section: avoiding {sorted(prefer_avoid_lab_rooms)}")
        
        # Schedule each course
        # Strict Priority Order: 1. Electives, 2. Combined Classes, 3. NML (Minor), 4. Regular Courses
        elective_courses = []
        combined_courses = []
        regular_courses = []
        
        for _, course in session_courses.iterrows():
            # Check if combined class
            combined_class_flag = False
            for colname in ['Combined Class', 'COMBINED CLASS', 'Combined Class ', 'COMBINED CLASS ']:
                if colname in course.index:
                    combined_class_val = str(course.get(colname, '')).strip().upper()
                    combined_class_flag = combined_class_val == 'YES'
                    break
            
            # Check if elective
            elective_flag = False
            for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                if colname in course.index:
                    elective_flag = str(course.get(colname, '')).upper() == 'YES'
                    if elective_flag:
                        break
            # Pattern overrides
            course_code_str = str(course.get('Course Code', '')).upper()
            course_name_str = str(course.get('Course Name', '')).upper()
            if 'ELEC' in course_code_str or 'ELEC' in course_name_str:
                elective_flag = True
            if 'HSS' in course_code_str or 'HSS' in course_name_str:
                elective_flag = False
            # Forced combined override (e.g., EC161)
            if course_code_str in FORCED_COMBINED_COURSES:
                combined_class_flag = True
            
            # AUTOMATIC HSS SYNCHRONIZATION:
            # All HSS courses MUST be combined across departments to avoid conflicts.
            if 'HSS' in course_code_str or 'HSS' in course_name_str:
                combined_class_flag = True
            
            if elective_flag:
                elective_courses.append(course)
            elif combined_class_flag:
                combined_courses.append(course)
            else:
                regular_courses.append(course)
        
        # Sort regular courses by total load (L+T+P) descending so heavier courses get first pick of slots
        def _course_load(c):
            l = int(pd.to_numeric(c.get('Lectures_Per_Week', 0), errors='coerce') or 0)
            t = int(pd.to_numeric(c.get('Tutorials_Per_Week', 0), errors='coerce') or 0)
            p = int(pd.to_numeric(c.get('Labs_Per_Week', 0), errors='coerce') or 0)
            return l + t + p
        regular_courses.sort(key=_course_load, reverse=True)
        
        # 1. Schedule Combined Classes first
        if combined_courses:
            print(f"  Scheduling {len(combined_courses)} combined courses...")
            for course in combined_courses:
                self._schedule_course(schedule, course, department, session, semester_id, prefer_avoid_lab_rooms=prefer_avoid_lab_rooms)
        
        # 2. Schedule Electives second
        if elective_courses:
            print(f"  Scheduling {len(elective_courses)} elective courses...")
            for course in elective_courses:
                self._schedule_course(schedule, course, department, session, semester_id, prefer_avoid_lab_rooms=prefer_avoid_lab_rooms)
        
        # 3. Schedule NML (Minor) third (early morning)
        self._schedule_minor_classes(schedule, department, session, semester_id)
        
        # 4. Schedule Regular Courses last
        if regular_courses:
            print(f"  Scheduling {len(regular_courses)} regular courses...")
            for course in regular_courses:
                self._schedule_course(schedule, course, department, session, semester_id, prefer_avoid_lab_rooms=prefer_avoid_lab_rooms)
        
        print(f"Schedule generated for {department} {session}")
        return schedule
    
    def get_actual_allocations(self, semester_id, department, session, course_code):
        """Get actual number of classes allocated for a course."""
        allocation_key = (semester_id, department, session, course_code)
        return self.actual_allocations.get(allocation_key, {
            'lectures': 0,
            'tutorials': 0,
            'labs': 0,
            'combined_class': False,
            'room': '',
            'lab_room': ''
        })
    
    def validate_room_conflicts(self):
        """Validate room allocation conflicts across all schedules.
        HARD RULE:
        - A physical classroom can host only ONE class at a time globally
          (across all semesters and departments).

        We still treat combined classes as one logical class when they share
        the same (course, session), even if multiple departments attend."""
        conflicts = []

        # Build a global view: (day, slot, room) -> list of (semester_key, dept, course, session)
        global_by_room_slot = {}
        for semester_key, semester_data in self.room_bookings.items():
            for (day, slot), bookings in semester_data.items():
                for b in bookings:
                    room = b.get('room', '')
                    if not room:
                        continue
                    key = (day, slot, room)
                    entry = (semester_key, b.get('dept'), b.get('course'), b.get('session'))
                    if key not in global_by_room_slot:
                        global_by_room_slot[key] = []
                    global_by_room_slot[key].append(entry)

        for (day, slot, room), entries in global_by_room_slot.items():
            # Check for conflicts within session buckets
            pre_courses = set()
            post_courses = set()
            
            for (sem_key, dept, course, session) in entries:
                c_code = str(course).strip()
                s_type = str(session).strip().lower()
                
                if "full" in s_type or "both" in s_type:
                    pre_courses.add(c_code)
                    post_courses.add(c_code)
                elif "pre" in s_type:
                    pre_courses.add(c_code)
                elif "post" in s_type:
                    post_courses.add(c_code)
                else:
                    # Treat unknown session as Full for safety (most restrictive)
                    pre_courses.add(c_code)
                    post_courses.add(c_code)

            # Conflict if more than one course in either bucket
            if len(pre_courses) > 1 or len(post_courses) > 1:
                conflicts.append({
                    'day': day,
                    'slot': slot,
                    'room': room,
                    'entries': entries,  # list of (semester_key, dept, course, session)
                })

        return conflicts

    def resolve_room_conflicts(self):
        """Attempt to automatically reallocate rooms so that there are ZERO clashes.
        - Keeps original timings fixed.
        - Moves extra classes in a conflicted (room, day, slot) to other free rooms.
        Returns the number of reallocated bookings."""
        conflicts = self.validate_room_conflicts()
        if not conflicts:
            return 0

        reallocated = 0

        # Helper: list of all room names we are allowed to use (non-lab + lab + C004 if present)
        all_rooms = []
        for name, _ in self.nonlab_rooms:
            if name and name not in all_rooms:
                all_rooms.append(name)
        for name, _ in self.lab_rooms:
            if name and name not in all_rooms:
                all_rooms.append(name)
        if self.c004_room:
            c004_name = self.c004_room[0]
            if c004_name and c004_name not in all_rooms:
                all_rooms.append(c004_name)

        for conflict in conflicts:
            day = conflict['day']
            slot = conflict['slot']
            room = conflict['room']
            # entries: list of (semester_key, dept, course, session)
            entries = conflict['entries']

            occ_key = (day, slot)
            if occ_key not in self.room_occupancy:
                continue
            used_rooms = self.room_occupancy[occ_key]

            # Keep the first logical class in the original room; move the rest
            for (semester_key, dept, course, session) in entries[1:]:
                slot_key = (day, slot)
                if semester_key not in self.room_bookings:
                    continue
                if slot_key not in self.room_bookings[semester_key]:
                    continue

                bookings_list = self.room_bookings[semester_key][slot_key]

                # Find this specific booking dict
                target_booking = None
                for b in bookings_list:
                    if (
                        b.get('room') == room and
                        b.get('dept') == dept and
                        str(b.get('course', '')).strip() == str(course).strip() and
                        b.get('session') == session
                    ):
                        target_booking = b
                        break
                if not target_booking:
                    continue

                # Find a new free room for this day/slot, respecting session
                new_room = None
                target_sess = str(session).strip().lower()
                for candidate in all_rooms:
                    if candidate == room:
                        continue
                    
                    # Room availability check (now session and course aware)
                    is_candidate_free = True
                    for r_name, r_sess, r_code in used_rooms:
                        if r_name == candidate:
                            # Same course can share
                            if str(r_code).strip().upper() == str(course).strip().upper():
                                continue
                            
                            existing_sess = str(r_sess).strip().lower()
                            if ("full" in existing_sess or "both" in existing_sess or
                                "full" in target_sess or "both" in target_sess or
                                ("pre" in target_sess and "pre" in existing_sess) or
                                ("post" in target_sess and "post" in existing_sess)):
                                is_candidate_free = False
                                break
                    
                    if is_candidate_free:
                        new_room = candidate
                        break

                if not new_room:
                    # No alternative room available at this time:
                    # drop this booking from the classroom allocation so the room stays clash-free.
                    # The class will still exist in the department timetable, but without a room.
                    print(f"WARNING: No free room for {course} ({dept}) at {semester_key} {day} {slot} – leaving without classroom.")
                    if target_booking in bookings_list:
                        bookings_list.remove(target_booking)
                    # Remove this room usage from occupancy
                    if room in used_rooms:
                        used_rooms.discard(room)
                    # Decrease allocation counter for this room
                    self.room_allocation_count[room] = max(
                        0, self.room_allocation_count.get(room, 1) - 1
                    )
                    # Also clear from assigned room lists
                    sem_id_str = semester_key.replace('sem_', '')
                    try:
                        sem_id = int(sem_id_str)
                    except ValueError:
                        sem_id = sem_id_str
                    allocation_key = (sem_id, dept, session, course)
                    if allocation_key in self.assigned_rooms:
                        self.assigned_rooms[allocation_key] = [
                            (d, s, r) for (d, s, r) in self.assigned_rooms[allocation_key]
                            if not (d == day and s == slot and r == room)
                        ]
                    if allocation_key in self.assigned_lab_rooms:
                        self.assigned_lab_rooms[allocation_key] = [
                            (d, s, r) for (d, s, r) in self.assigned_lab_rooms[allocation_key]
                            if not (d == day and s == slot and r == room)
                        ]
                    # Count as reallocation (handled conflict by dropping room)
                    reallocated += 1
                    continue

                # Update booking record
                target_booking['room'] = new_room

                # Update occupancy: now store course_code too
                used_rooms.add((new_room, session, course))

                # Adjust allocation counters
                self.room_allocation_count[room] = self.room_allocation_count.get(room, 1) - 1
                self.room_allocation_count[new_room] = self.room_allocation_count.get(new_room, 0) + 1

                # Update assigned room lists so summaries/faculty view stay consistent
                sem_id_str = semester_key.replace('sem_', '')
                try:
                    sem_id = int(sem_id_str)
                except ValueError:
                    sem_id = sem_id_str

                allocation_key = (sem_id, dept, session, course)
                updated = False

                if allocation_key in self.assigned_rooms:
                    for idx, (d, s, r) in enumerate(self.assigned_rooms[allocation_key]):
                        if d == day and s == slot and r == room:
                            self.assigned_rooms[allocation_key][idx] = (d, s, new_room)
                            updated = True
                            break

                if not updated and allocation_key in self.assigned_lab_rooms:
                    for idx, (d, s, r) in enumerate(self.assigned_lab_rooms[allocation_key]):
                        if d == day and s == slot and r == room:
                            self.assigned_lab_rooms[allocation_key][idx] = (d, s, new_room)
                            updated = True
                            break

                reallocated += 1

        # After reallocation, room_bookings / room_occupancy reflect clash-free allocation (as far as possible)
        return reallocated