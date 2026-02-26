import pandas as pd
import os
import glob

def check_conflicts():
    print("Checking for conflicts in generated timetables...")
    
    # Get all semX_timetable.xlsx files from the output directory
    output_dir = "output"
    timetable_files = glob.glob(os.path.join(output_dir, "sem*_timetable.xlsx"))
    if not timetable_files:
        print("No timetable files found. Please run main.py first.")
        return

    all_conflicts = []

    for file_path in timetable_files:
        print(f"\nAnalyzing {file_path}...")
        try:
            xl = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

        # We care about CSE-A and CSE-B primarily, but check all dept sheets
        dept_sheets = [s for s in xl.sheet_names if s not in ['Allocation Summary', 'Course Details', 'Electives', 'Minors']]
        
        # Structure to hold assigned slots: (semester, day, slot) -> list of (dept, activity)
        global_schedule = {}

        for sheet in dept_sheets:
            df = xl.parse(sheet)
            # Find the grid starting point (usually rows are times, columns are days)
            # Or identify the layout. Assuming a standard grid based on ExcelExporter.
            
            # For simplicity in this validator, let's look at 'Allocation Summary' if possible 
            # for exact slot data, or parse the grids. 
            # Parsing the grids is more robust for checking what students actually see.
            
            # Let's assume the grid has 'Day' or similar logic
            # Actually, the 'Allocation Summary' contains (Course, Dept, Day, Slot, Room)
            pass

        # Robust approach: Load the 'Allocation Summary' if it exists in any file
        # or load from the primary output if main.py produces one master file.
        # Based on excel_loader and exporter, each sem has its own file.
        
        summary_sheet = 'Allocation Summary'
        if summary_sheet in xl.sheet_names:
            df_summary = xl.parse(summary_sheet)
            # print allocation percentage if available
            try:
                pct = df_summary.loc[df_summary['Metric']=='Allocation %','Value'].iloc[0]
                print(f"  Allocation percentage: {pct}")
                if isinstance(pct, str) and pct.strip().endswith('%'):
                    try:
                        val = float(pct.strip().strip('%'))
                        if val < 100:
                            print(f"  WARNING: allocation percentage below 100% in {file_path}")
                    except ValueError:
                        pass
            except Exception:
                pass
            # show failures sheet if present
            if 'Allocation_Failures' in xl.sheet_names:
                df_fail = xl.parse('Allocation_Failures')
                if not df_fail.empty:
                    print('  Unallocated courses:')
                    print(df_fail.head(20))
            # Required columns: 'Department', 'Day', 'Time Slot', 'Course Code', 'Type'
            # Note: excel_exporter uses 'Time Slot'/ 'Day' etc.
            
            # Group by Day, Time Slot, and Department
            # A student group (Dept) should have at most one activity at a time.
            
            # 1. Student-Level Conflict: (Day, Slot, Dept) must be unique
            clashes = df_summary[df_summary.duplicated(subset=['Day', 'Time Slot', 'Department'], keep=False)]
            if not clashes.empty:
                for _, row in clashes.iterrows():
                    all_conflicts.append(f"Student Clash: Dept {row['Department']} at {row['Day']} {row['Time Slot']} - Multiple entries.")

            # 2. Elective/Combined Blocking:
            # If any Dept has 'Elective' or 'Elective Tut', no other core class for any dept in same sem?
            # User said: "3-ELEC and 3-ELEC Tutorial must be treated as blocking slots for ALL students."
            
            for (day, slot), group in df_summary.groupby(['Day', 'Time Slot']):
                has_elective = any(str(row['Course Code']).upper().find('ELEC') != -1 or str(row.get('Type', '')).upper() == 'ELECTIVE' for _, row in group.iterrows())
                has_combined = any(str(row.get('Combined', '')).upper() == 'YES' or str(row.get('Type', '')).upper() == 'COMBINED' for _, row in group.iterrows())
                
                if has_elective:
                    # Check if there are any non-elective classes in this slot
                    non_electives = group[~(group['Course Code'].astype(str).str.upper().str.contains('ELEC') | (group.get('Type', '').astype(str).str.upper() == 'ELECTIVE'))]
                    # Filter out Minor (if allowed parallel? No, "blocking slots for ALL students")
                    # Filter out HSS (Elective constraint says "No core/common class")
                    if not non_electives.empty:
                        for _, row in non_electives.iterrows():
                            all_conflicts.append(f"Elective Block Violation: Dept {row['Department']} has {row['Course Code']} during Elective slot at {day} {slot}")

                if has_combined:
                    # Combined classes should be at the same time for both sections.
                    # No elective parallel to combined.
                    pass

    if not all_conflicts:
        print("\nSUCCESS: No conflicts found across all semesters!")
    else:
        print(f"\nFAILURE: Found {len(all_conflicts)} conflicts:")
        for conflict in all_conflicts[:20]: # Show first 20
            print(f"  - {conflict}")
        if len(all_conflicts) > 20:
            print(f"  ... and {len(all_conflicts) - 20} more.")

if __name__ == "__main__":
    check_conflicts()
