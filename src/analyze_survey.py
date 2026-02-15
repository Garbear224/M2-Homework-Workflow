import pandas as pd
import matplotlib.pyplot as plt
import os

def load_data(filepath):
    """Loads the dataset."""
    print(f"Loading data from {filepath}...")
    try:
        # Read CSV with header=None to get raw rows
        df = pd.read_csv(filepath, header=None)
        return df
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

def map_group_to_score(val):
    """Maps group names or numeric values to scores."""
    if pd.isna(val):
        return None

    # Check if value is numeric
    try:
        return float(val)
    except ValueError:
        pass

    # Check if value is string and map
    s_val = str(val)
    if "Most Beneficial" in s_val:
        return 3
    elif "Neutral" in s_val:
        return 2
    elif "Least Beneficial" in s_val:
        return 1

    return None

def main():
    filepath = 'data/Grade Program Exit Survey Data.csv'

    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return

    df_raw = load_data(filepath)
    if df_raw is None:
        return

    # Row 0: Question IDs (e.g. Q35_1)
    # Row 1: Course Names
    qid_row = df_raw.iloc[0]
    course_row = df_raw.iloc[1]

    # Data starts from row 2
    data = df_raw.iloc[2:]

    course_ratings = []

    print("\nMapping Question IDs to Course Names:")
    print("-" * 60)
    print(f"{'Question ID':<15} | {'Course Name'}")
    print("-" * 60)

    for col_idx in range(len(df_raw.columns)):
        qid = str(qid_row[col_idx])
        course = str(course_row[col_idx])

        # Check if QID starts with Q35 (Core) or Q76 (Elective)
        if qid.startswith('Q35') or qid.startswith('Q76'):
            print(f"{qid:<15} | {course}")

            col_data = data[col_idx]

            for idx, val in col_data.items():
                score = map_group_to_score(val)
                if score is not None:
                     course_ratings.append({
                        'Respondent': idx,
                        'Course': course,
                        'Score': score
                    })

    print("-" * 60)

    ratings_df = pd.DataFrame(course_ratings)

    if ratings_df.empty:
        print("No ratings found. Check data extraction logic or ensure data file is correct.")
        return

    # Calculate Average Rating
    course_stats = ratings_df.groupby('Course')['Score'].agg(['mean', 'count']).reset_index()
    course_stats.columns = ['Course', 'Average Rating', 'Count']

    # Sort
    course_stats = course_stats.sort_values('Average Rating', ascending=False)

    print("\nCourse Rankings:")
    print(course_stats)

    # Plot
    plt.figure(figsize=(12, 8))
    plt.barh(course_stats['Course'], course_stats['Average Rating'], color='skyblue')
    plt.xlabel('Average Rating')
    plt.title('Course Rankings')
    plt.gca().invert_yaxis()
    plt.tight_layout()

    output_path = 'outputs/rank_order.png'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    plt.savefig(output_path)
    print(f"\nChart saved to {output_path}")

if __name__ == "__main__":
    main()
