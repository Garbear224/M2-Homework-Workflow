import pandas as pd
import matplotlib.pyplot as plt
import os
import re

def analyze_survey():
    # File path
    file_path = 'Grad Program Exit Survey Data.xlsx'

    # Read the Excel file
    # header=[0, 1] reads the first two rows as headers
    # Row 0 (index 0) contains the question text (e.g. "Please identify which MAcc CORE courses...")
    # Row 1 (index 1) contains the ImportId (e.g. "{"ImportId":"QID84_G0_1_RANK"}")
    try:
        df = pd.read_excel(file_path, header=[0, 1])
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return

    # Helper function to extract info from header
    def parse_header(header_tuple):
        # header_tuple[0] is the question text
        # header_tuple[1] is the ImportId
        q_text = str(header_tuple[0])

        # Check if it's a relevant question (Core or Elective)
        is_core = "CORE" in q_text
        is_elective = "Elective" in q_text

        if not (is_core or is_elective):
            return None

        # Extract Course Name and Group
        # Pattern: ... - Ranks - [Group] - [Course Name] - Rank
        # Note: The separator is usually " - "
        match = re.search(r' - Ranks - (.*?) - (.*?) - Rank', q_text)
        if match:
            group = match.group(1)
            course_name = match.group(2)
            return {
                'type': 'Core' if is_core else 'Elective',
                'group': group,
                'course': course_name,
                'col_idx': header_tuple
            }
        return None

    # Identify relevant columns
    relevant_columns = []
    for col in df.columns:
        info = parse_header(col)
        if info:
            relevant_columns.append(info)

    if not relevant_columns:
        print("No relevant columns found (looking for 'CORE' or 'Elective' in header).")
        return

    print(f"Found {len(relevant_columns)} relevant columns.")

    # Audit Trail: Map Question IDs (from ImportId or implied) to Course Names
    print("\nAudit Trail: Mapping Question IDs to Course Names")
    # We'll map the ImportId (e.g. QID84_...) to the Course Name
    # Or just list the unique courses found
    unique_courses = {}
    for info in relevant_columns:
        # ImportId is in info['col_idx'][1]
        import_id = info['col_idx'][1]
        course_name = info['course']
        if course_name not in unique_courses:
             unique_courses[course_name] = []
        unique_courses[course_name].append(import_id)

    for course, ids in unique_courses.items():
        # Just show one ID as example or count
        print(f"Course: {course} (IDs: {ids[0]}...)")

    # Process Data
    # We need to calculate a score for each course for each respondent
    # Scoring:
    # Most Beneficial: 3
    # Neutral: 2
    # Least Beneficial: 1
    # Did not take: NaN (ignore)

    # We will reshape the data to long format or iterate row by row
    # Since we need average per course, we can process by course

    course_scores = {} # course_name -> list of scores

    # Group columns by course
    courses = list(unique_courses.keys())

    for course in courses:
        # Find columns for this course
        cols = [info for info in relevant_columns if info['course'] == course]

        # Extract data for these columns
        # We need to find which column has a value for each row
        # Since a respondent can only put a course in one group (usually), we take the max score?
        # Wait, if they put it in multiple (which shouldn't happen), we need to handle it.
        # But let's assume valid data.

        # Create a temporary dataframe for this course
        # Map columns to scores
        course_df = pd.DataFrame()

        for info in cols:
            col_name = info['col_idx']
            group = info['group']

            score = None
            if group == 'Most Beneficial':
                score = 3
            elif group == 'Neutral':
                score = 2
            elif group == 'Least Beneficial':
                score = 1
            elif group == 'Did not take':
                score = None # Ignore

            if score is not None:
                # If the cell has a value (rank), it means the user put it in this group
                # We are interested in the presence of a value, not the rank itself (based on our scoring model)
                # Ensure we handle NaN correctly
                series = df[col_name].notna().astype(int) * score
                # Replace 0 with NaN if we want to ignore non-selected, but here 0 means "not in this group"
                # We want to know which group it IS in.
                # So if series is 3, it's in Most Beneficial. If 0, it's not.
                course_df[group] = series.replace(0, float('nan'))

        # Now for each row, we expect only one column to have a non-NaN value
        # We take the max (or mean) across the row to get the single score for that respondent
        if not course_df.empty:
            respondent_scores = course_df.max(axis=1) # Returns NaN if all are NaN
            # Filter out NaNs (Did not take or didn't answer)
            valid_scores = respondent_scores.dropna()
            course_scores[course] = valid_scores.mean()
        else:
            course_scores[course] = float('nan')

    # Create Series
    results_series = pd.Series(course_scores)

    # Rank from highest to lowest
    results_sorted = results_series.sort_values(ascending=False)

    print("\nCourse Rankings (Highest to Lowest Average Rating):")
    print(results_sorted)

    # Ensure outputs directory exists
    output_dir = 'outputs'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Visualization: Horizontal bar chart
    plt.figure(figsize=(12, 10)) # Increased size for better readability
    results_sorted.plot(kind='barh', color='skyblue')
    plt.xlabel('Average Rating (3=Most Beneficial, 2=Neutral, 1=Least Beneficial)')
    plt.title('Course Ratings (Highest to Lowest)')
    plt.gca().invert_yaxis() # To have the highest rating at the top
    plt.tight_layout()

    output_file = os.path.join(output_dir, 'rank_order.png')
    plt.savefig(output_file)
    print(f"\nChart saved to {output_file}")

if __name__ == "__main__":
    analyze_survey()
