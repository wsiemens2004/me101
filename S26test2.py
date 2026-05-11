import pandas as pd
from collections import defaultdict

# Load input Excel file
input_filename = 'Test4.xlsx'
df = pd.read_excel(input_filename)

print(df.columns)

# Column names for group and member number
#WARNING: any edits to the questions or column titles need to be edited here as well
group_id_col = 'What is your group number? (see F25 Peer Evaluation Group Member Numbers ) (1-18)'
member_num_col = 'What is your group member number? (see F25 Peer Evaluation Group Member Numbers ) (remember you will be evaluating yourself as well)'
if group_id_col not in df.columns or member_num_col not in df.columns:
    raise ValueError("Group ID or Member Number columns not found in the input file")

# Step 1: Detect max group members (from column names)
max_members = 0
for col in df.columns:
    #print('col')
    #print(col)
    if "Group Member #" in col and "What is the full name" in col:
        try:
            num = int(col.split('#')[1].split()[0])
            #print('num')
            #print(num)
            
            if num > max_members:
                max_members = num
                #print('max_members')
                #print(max_members)
                
        except:
            pass
print('max_members')
print(max_members)
# Step 2: Build canonical name map for each (group_id, member_num)
# Use the first seen name for that pair
group_member_map = {}

for _, row in df.iterrows():
    group_id = row[group_id_col]
    print('group_id')
    print(group_id)
    member_num = row[member_num_col]
    print('member_num')
    print(member_num)
    
    # member_num might be numeric or string, convert to int safely
    try:
        i = int(member_num)
        print('i')
        print(i)
    except:
        print('fail')
        continue

    if i == 1:
        name_col = 'What is the full name of Group Member #1 in your group?'
    elif i == 2:
        name_col = 'What is the full name of Group Member #2 in your group?'
    elif i == 3:
        name_col = 'What is the full name of Group Member #3 in your group?'
    elif i == 4:
        name_col = 'What is the full name of Group Member #4 in your group?'
    elif i == 5:
        name_col = 'What is the full name of Group Member #5 in your group?'
    elif i == 6:
        name_col = 'What is the full name of Group Member #6 in your group? (or write "none" if no 6th group member)'
    else:
        continue
    print('name col')
    print(name_col)

    if name_col not in df.columns:
        continue

    name_raw = str(row.get(name_col, None))
    print('name_raw')
    print(name_raw)
    print(pd.notna(name_raw))
    
    if pd.notna(name_raw):
        name = name_raw.strip()
        key = (group_id, i)
        if key not in group_member_map:
            group_member_map[key] = name

# Step 3: Collect all evaluations for each (group_id, member_num)
evaluatee_data = defaultdict(list)

for _, row in df.iterrows():
    evaluator_email = row.get('Email Address', None)
    print('evaluator_email')
    print(evaluator_email)
    group_id = row[group_id_col]
    print('group_id')
    print(group_id)
    
    for i in range(1, max_members + 1):
        if i == 1:
            name_col = 'What is the full name of Group Member #1 in your group?'
        elif i == 2:
            name_col = 'What is the full name of Group Member #2 in your group?'
        elif i == 3:
            name_col = 'What is the full name of Group Member #3 in your group?'
        elif i == 4:
            name_col = 'What is the full name of Group Member #4 in your group?'
        elif i == 5:
            name_col = 'What is the full name of Group Member #5 in your group?'
        elif i == 6:
            name_col = 'What is the full name of Group Member #6 in your group? (or write "none" if no 6th group member)'
        else:
            continue

        if name_col not in df.columns:
            continue

        name_raw = str(row.get(name_col, None))
        if pd.isna(name_raw):
            continue
        name = name_raw.strip()

        # Find the canonical name by group_id and member number
        key = (group_id, i)
        canonical_name = group_member_map.get(key)
        if not canonical_name:
            # fallback to current name if missing in map (should not happen)
            canonical_name = name
            canonical_i = i

        # Extract ratings for this group member
        entry = {
            
            'Evaluator Email': evaluator_email,
            'Evaluatee Name (as entered)': name,            
            'Problem Solving': row.get(
                f'Your rating of Group Member #{i} on Problem solving: How much did the team member contribute to creatively solving and documenting solutions to the problems? (Note: answer all questions for all group members, including yourself.)',
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
                f'What would you say was Group Member #{i}\'s most valuable contribution to the project, labs, or the functioning of the group?',
                None),
            'Weakness': row.get(
                f'Your comments on Group Member #{i}\'s weakness(es). Anything else you would like to say about working with this group member?',
                None)
        }
        evaluatee_data[canonical_name].append(entry)
    
print('evaluatee_data')
print(evaluatee_data)

#INSTRUCTIONS: be careful in the input file.... it matters what people call themselves, so if they got their member number wrong or they wrote "none" that needs to be fixed
#this means you have to add people who didnt fill out the form and put their name in the member # for their row

# Step 4: Build output rows for every group member (use canonical names)
output_rows = []

for (group_id, member_num), canonical_name in sorted(group_member_map.items()):
    evaluations = evaluatee_data.get(canonical_name, [])

    new_row = {
        'Evaluatee': canonical_name,
        'Group ID': group_id,
        'Member #': member_num,
    }

    #  FORCE these columns to exist even if some evaluators are missing
    for j in range(1, max_members + 1):
        new_row[f'Evaluatee Name (as entered) {j}'] = None

    # Fill them in aligned with evaluator index (1..N)
    for idx, eval_dict in enumerate(evaluations, start=1):
        new_row[f'Evaluatee Name (as entered) {idx}'] = eval_dict.get('Evaluatee Name (as entered)')

        for k, v in eval_dict.items():
            if k == 'Evaluatee Name (as entered)':
                continue
            new_row[f'{k} {idx}'] = v

    output_rows.append(new_row)

output_df = pd.DataFrame(output_rows)

# Safe reorder (won't KeyError)
name_cols = [f'Evaluatee Name (as entered) {j}' for j in range(1, max_members + 1)]
front_wanted = ['Evaluatee'] + name_cols + ['Group ID', 'Member #']
front = [c for c in front_wanted if c in output_df.columns]
rest = [c for c in output_df.columns if c not in front]
output_df = output_df[front + rest]


# keep whatever other columns exist after that
rest = [c for c in output_df.columns if c not in front]
output_df = output_df[front + rest]




# Save to Excel
output_filename = 'aggregated_peer_evaluations_fixed.xlsx'
output_df.to_excel(output_filename, index=False)

print(f"Aggregated peer evaluations saved to {output_filename}")

