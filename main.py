import pandas as pd
import re
import os
# Load the excel file
file_path = 'Daily Report.xlsx'
if not os.path.exists(file_path):
    # Try to find any xlsx file if the name is slightly different
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    if files:
        file_path = files[0]
    else:
        raise FileNotFoundError("No Excel file found in workspace.")
 print(f"Processing file: {file_path}")
_# Read columns H and M (0-indexed: 7 and 12)
# We assume the first row is header. If not, adjust accordingly.
try:
    df = pd.read_excel(file_path)
except Exception as e:
    # Try reading without header if standard read fails
    df = pd.read_excel(file_path, header=None)
    # Assign dummy headers to access by index later if needed, but let's assume standard first
    # Re-try with header=None and manually pick columns 7 and 12
    df = pd.read_excel(file_path, header=None)
    col_h = df.iloc[:, 7]
    col_m = df.iloc[:, 12]
    raw_tags = pd.concat([col_h, col_m], ignore_index=True).dropna().astype(str)
if 'H' in df.columns and 'M' in df.columns:
    raw_tags = pd.concat([df['H'], df['M']], ignore_index=True).dropna().astype(str)
elif df.shape[1] > 12:
    # Fallback to index based if headers are missing or different
    raw_tags = pd.concat([df.iloc[:, 7], df.iloc[:, 12]], ignore_index=True).dropna().astype(str)
else:
    # If columns don't exist, try to find them by content or error out
    # Assuming standard structure based on previous turns
    raise ValueError("Could not identify columns H and M. Please check file structure.")
# Clean initial data
raw_tags = raw_tags.str.strip()
raw_tags = raw_tags[raw_tags != 'nan']
raw_tags = raw_tags[raw_tags != '']
raw_tags = raw_tags[raw_tags.str.lower() != 'done']
# Filter out dates (simple heuristic: contains / or - and numbers only mostly)
# But be careful not to filter valid tags with numbers.
# The user said "tags like dates and DONE were filtered".
# We will rely on the specific tag patterns provided in the prompt logic later.
processed_tags = []
def normalize_tag(tag):
    """Normalize tag: remove spaces, handle hyphens consistently."""
    tag = str(tag).strip()
    # Remove extra spaces
    tag = re.sub(r'\s+', ' ', tag).strip()
    return tag
def expand_range(tag_str):
    """Expand ranges like GD-0009/0021 -> GD-0009, GD-0010, ... GD-0021"""
    tag_str = normalize_tag(tag_str)
    # Pattern for Range: PREFIX-NUMBER/NUMBER (e.g., GD-0009/0021)
    # Also handles GD-9/21 potentially
    range_match = re.match(r'^([A-Za-z]+)-?(\d+)/(\d+)$', tag_str.replace(' ', ''))
    if range_match:
        prefix = range_match.group(1)
        start = int(range_match.group(2))
        end = int(range_match.group(3))
        expanded = []
        # Determine padding based on the start number string length or a standard (e.g., 4 digits)
        # Looking at example GD-0009, it seems 4 digits padding is safe, or match the input style.
        # Let's try to preserve the width of the start number if possible, or default to 4 for industrial tags.
        width = len(range_match.group(2))
        if width < 2: width = 4 # Default to 4 if single digit like 9
        for i in range(start, end + 1):
            expanded.append(f"{prefix}-{str(i).zfill(width)}")
        return expanded
    return [tag_str]
def resolve_conflicts(tags_list):
    """
    1. Normalize hyphens (treat 'P 2902' and 'P-2902' as same).
    2. Priority: GA > P. If both exist, keep GA.
    """
    # Map normalized key -> best version
    tag_map = {}
    for tag in tags_list:
        # Normalize for comparison: remove hyphens and spaces, lowercase
        key = re.sub(r'[-\s]', '', tag).lower()
        current_best = tag_map.get(key)
        if current_best is None:
            tag_map[key] = tag
        else:
            # Check priority: GA vs P
            # We need to check the original prefixes
            tag_upper = tag.upper()
            best_upper = current_best.upper()
            # Identify prefix
            def get_prefix(t):
                m = re.match(r'^([A-Za-z]+)', t.replace('-', '').replace(' ', ''))
                return m.group(1) if m else ""
            p_tag = get_prefix(tag_upper)
            p_best = get_prefix(best_upper)
            # Rule: GA has priority over P
            if p_tag == 'GA' and p_best == 'P':
                tag_map[key] = tag
            elif p_tag == 'P' and p_best == 'GA':
                pass # Keep existing GA
            else:
                # If same prefix or other, keep the one that looks more standard (with hyphen usually)
                # Or just keep the first one found if no conflict rule applies.
                # User said: "if tag with - and without - exists, know they are same and separate"
                # This implies deduplication. We keep one.
                # Let's prefer the one with hyphen as it's standard industrial format
                if '-' in tag and '-' not in current_best:
                    tag_map[key] = tag
                elif '-' not in tag and '-' in current_best:
                    pass
                else:
                    pass # Keep first
    return list(tag_map.values())
# Step 1: Collect all raw tags
all_raw = raw_tags.tolist()
# Step 2: Expand ranges
expanded_tags = []
for t in all_raw:
    expanded_tags.extend(expand_range(t))
# Step 3: Resolve conflicts and deduplicate
final_tags = resolve_conflicts(expanded_tags)
# Sort for neatness
final_tags.sort(key=lambda x: (re.sub(r'\d+', '', x).lower(), [int(c) for c in re.findall(r'\d+', x)]))
print(f"Total unique tags after processing: {len(final_tags)}")
print("Sample tags:", final_tags[:10])
# Save to tag.xlsx
df_output = pd.DataFrame(final_tags, columns=['Tag'])
df_output.to_excel('tag.xlsx', index=False)
print("Saved tag.xlsx")
# Update Daily Report.xlsx: Add/Replace 'Tags' sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_output.to_excel(writer, sheet_name='Tags', index=False)
print(f"Updated {file_path} with 'Tags' sheet.")
print("Done!")
EOF
OpenBLAS WARNING - could not determine the L2 cache size on this system, assuming 256k
Processing file: Daily Report.xlsx
Total unique tags after processing: 32459
Sample tags: ['0', '111111111111111001872793600', '11111111111111100471131876275277446358205726720', '2603 با 2602 ju,dq an.', '2 بار پایش و تغییر بار برای 5105D انجام گردید بار در 95 درصد قرار داده شد 3 بار پایش و تغییر بار برای 5105C انجام گردید بار در 95 درصد قرار داده شد', '2 تا از فیوزها مربوط به سوکت های جنوب مخازن تعویض شد و سیم کشی ان اصلاح شد', '8 تا از فیکسچرها مشکل روشنایی داشت که بعد از تعویض لامپ های سوخته مشکل برطرف شد', '5 تا از لامپ ها تعویض گردید', '4 تا سلونویید ولو لوپ زده شد', '5 تا موتور فیدر و الکتروموتور ان اماده FREE RUN شد']
