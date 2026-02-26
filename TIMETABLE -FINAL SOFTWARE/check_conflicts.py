import sys
import os
import io

# Force UTF-8 regarding stdout for any print calls
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.append(os.getcwd())

try:
    from main import TimetableGenerator
except ImportError as e:
    with open("result.txt", "w", encoding='utf-8') as f:
        f.write(f"Import Error: {e}")
    sys.exit(1)

def run_check():
    with open("result.txt", "w", encoding='utf-8') as f:
        f.write("Initializing generator...\n")
        try:
            generator = TimetableGenerator()
            generator.setup_environment()
            
            f.write("Running generation (this might take a moment)...\n")
            # This generates data structures needed for validation
            # It will also write excel files, which is fine/unavoidable easily without deep code changes
            generator.generate_timetables() 
            
            f.write("Validating room conflicts...\n")
            conflicts = generator.schedule_generator.validate_room_conflicts()
            
            if conflicts:
                f.write(f"\nFOUND {len(conflicts)} CONFLICTS:\n")
                found_conflicts = []
                for c in conflicts:
                    sem = c.get('semester', 'Unknown Sem')
                    day = c.get('day', 'Unknown Day')
                    slot = c.get('slot', 'Unknown Slot')
                    room = c.get('room', 'Unknown Room')
                    entries = c.get('entries', []) # list of (semester_key, dept, course, session)
                    
                    entry_strs = []
                    for entry in entries:
                        # entries is a list of tuples (semester_key, dept, course, session)
                        dept = entry[1]
                        course = entry[2]
                        session_type = entry[3]
                        entry_strs.append(f"{dept}:{course}({session_type})")
                    
                    conflict_str = f"CONFLICT: {sem} {day} {slot} in {room} -> {', '.join(entry_strs)}"
                    f.write(conflict_str + "\n")
                    found_conflicts.append(conflict_str)
                print(f"Conflicts found: {len(found_conflicts)}")
            else:
                f.write("\nNO ROOM CONFLICTS FOUND.\n")
                print("No room conflicts found.")
                
        except Exception as e:
            f.write(f"Error during check: {e}\n")
            import traceback
            traceback.print_exc(file=f)
            print(f"Error: {e}")

if __name__ == "__main__":
    run_check()
