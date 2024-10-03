import pandas as pd

# Load the data from Excel files
sign_in = pd.read_excel("2024企業實習成果發表會報名表單 (回覆).xlsx")
vote = pd.read_excel("2024企業實習發表會人氣獎表單 (回覆).xlsx")
group_nums = 14

# Extract relevant columns
sign_in_ids = sign_in.iloc[:, 3]  # D column for student ID in the sign-in form
check_box = sign_in.iloc[:, 7]  # H column for checkbox in the sign-in form (True/False)
vote_ids = vote.iloc[:, 2]  # C column for student ID in the vote form
vote_columns = vote.iloc[:, 6:9]  # G, H, and I columns representing 3 votes for groups

# Combine vote_ids and vote_columns for easier processing
vote_data = pd.concat([vote_ids, vote_columns], axis=1)

# Keep only the last submission for each student ID
vote_data_last = vote_data.drop_duplicates(subset=vote_data.columns[0], keep='last')

# Create a dictionary to store vote counts for each group (assuming 14 groups)
group_votes = {f"Group {i}": 0 for i in range(1, group_nums + 1)}

# Iterate through each row in the last submissions of the vote form
for i, row in vote_data_last.iterrows():
    student_id = row.iloc[0]  # Student ID (C column)

    # Check if the student ID exists in the sign-in form and the checkbox is True (H column)
    if student_id in sign_in_ids.values:
        sign_in_index = sign_in_ids[sign_in_ids == student_id].index[0]
        if check_box[sign_in_index]:  # Only count votes if the checkbox is True
            # Get unique groups voted for by this student (ensuring 1 vote per group)
            unique_groups = set()
            for vote in row.iloc[1:]:  # Process votes in G, H, I columns
                if pd.notna(vote):  # Check if the vote is not NaN
                    unique_groups.add(int(vote))  # Add the group number to the set

            # Add 1 vote to each group (1 vote per group, even if multiple are given)
            for group_num in unique_groups:
                if group_num in range(1, group_nums + 1):  # Ensure the group number is valid (1 to 14)
                    group_votes[f"Group {group_num}"] += 1

# Print the votes for each group
print("Votes for all groups:")
for group, votes in group_votes.items():
    print(f"{group}: {votes} votes")
