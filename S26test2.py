bimport pandas as pd
from collections import defaultdict

# Load input Excel file
input_filename = 'me101sp25evalsfinal(1).xlsx'
df = pd.read_excel(input_filename)

# Column names for group and member number
group_id_col = 'What is your group number?'
member_num_col = 'What is your group member number? (again, see Peer Evaluation Group Member Numbers and keep it handy for the rest of this survey)'

if group_id_col not in df.columns or member_num_col not in df.columns:
    raise ValueError("Group ID or Member Number columns not found in the input file")

# Step 1: Detect max group members (from column names)
max_members = 0
for col in df.columns:
    if "Group Member #" in col and "What is the full name" in col:
        try:
            num = int(col.split('#')[1].split()[0])
            if num > max_members:
                max_members = num
        except:
            pass

# Step 2: Build canonical name map for each (group_id, member_num)
# Use the first seen name for that pair
group_member_map = {}

for _, row in df.iterrows():
    group_id = row[group_id_col]
    member_num = row[member_num_col]

    # member_num might be numeric or string, convert to int safely
    try:
        i = int(member_num)
    except:
        continue

    if i == 1:
        name_col = 'What is the full name of Group Member #1 (remember to use our numbering!) in your group?'
    elif i == 2:
        name_col = 'What is the full name of Group Member #2 in your group?'
    elif i == 3:
        name_col = 'What is the full name of Group Member #3 in your group?'
    elif i == 4:
        name_col = 'What is the full name of Group Member #4 in your group?'
    elif i == 5:
        name_col = 'What is the full name of Group Member #5 in your group?'
    elif i == 6:
        name_col = 'What is the full name of Group Member #6 in your group?'
    else:
        continue

    if name_col not in df.columns:
        continue

    name_raw = row.get(name_col, None)
    if pd.notna(name_raw):
        name = name_raw.strip()
        key = (group_id, i)
        if key not in group_member_map:
            group_member_map[key] = name

# Step 3: Collect all evaluations for each (group_id, member_num)
evaluatee_data = defaultdict(list)

for _, row in df.iterrows():
    evaluator_email = row.get('Email Address', None)
    group_id = row[group_id_col]

    for i in range(1, max_members + 1):
        if i == 1:
            name_col = 'What is the full name of Group Member #1 (remember to use our numbering!) in your group?'
        elif i == 2:
            name_col = 'What is the full name of Group Member #2 in your group?'
        elif i == 3:
            name_col = 'What is the full name of Group Member #3 in your group?'
        elif i == 4:
            name_col = 'What is the full name of Group Member #4 in your group?'
        elif i == 5:
            name_col = 'What is the full name of Group Member #5 in your group?'
        elif i == 6:
            name_col = 'What is the full name of Group Member #6 in your group?'
        else:
            continue

        if name_col not in df.columns:
            continue

        name_raw = row.get(name_col, None)
        if pd.isna(name_raw):
            continue
        name = name_raw.strip()

        # Find the canonical name by group_id and member number
        key = (group_id, i)
        canonical_name = group_member_map.get(key)
        if not canonical_name:
            # fallback to current name if missing in map (should not happen)
            canonical_name = name

        # Extract ratings for this group member
        entry = {
            'Evaluator Email': evaluator_email,
            'Problem Solving': row.get(
                f'Your rating of Group Member #{i} on Problem solving: How much did the team member contribute to creatively defining, solving, and documenting solutions to the problem? (Note: answer all questions for all group members, including yourself.)',
                None),
            'Effort': row.get(
                f'Your rating of Group Member #{i} on Effort: Did the team member do his/her share of the work?', None),
            'Reliability': row.get(
                f'Your rating of Group Member #{i} on Reliability: Did the team member communicate and perform work reliably, completely, and on time?',
                None),
            'Team Support': row.get(
                f'Your rating of Group Member #{i} on Team Support: Did the team member contribute by attitude and action to team morale?',
                None),
            'Overall': row.get(
                f'Your rating of Group Member #{i} Overall: How effective was this team member? How valuable was his/her contribution?',
                None),
            'Work Again': row.get(f'Would you choose to work with this team member (#{i}) again?', None),
            'Strength': row.get(
                f'What would you say was Group Member #{i}\'s most valuable contribution to the project or the functioning of the group?',
                None),
            'Weakness': row.get(
                f'Your comments on Group Member #{i}\'s weakness(es). Anything else you would like to say about working with this group member?',
                None)
        }
        evaluatee_data[canonical_name].append(entry)

# Step 4: Build output rows for every group member (use canonical names)
output_rows = []
for canonical_name in sorted(group_member_map.values()):
    evaluations = evaluatee_data.get(canonical_name, [])
    row = {'Evaluatee': canonical_name}
    for idx, eval_dict in enumerate(evaluations, start=1):
        for k, v in eval_dict.items():
            row[f'{k} {idx}'] = v
    output_rows.append(row)

output_df = pd.DataFrame(output_rows)

# Save to Excel
output_filename = 'aggregated_peer_evaluations_fixed.xlsx'
output_df.to_excel(output_filename, index=False)

print(f"Aggregated peer evaluations saved to {output_filename}")
