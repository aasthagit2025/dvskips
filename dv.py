import pandas as pd
import re

# --- Step 1: Convert Skip Excel to Validation Rules ---
def convert_skip_excel(skip_file):
    df = pd.read_excel(skip_file)
    rules = []

    for _, row in df.iterrows():
        skip_from = str(row.get("Skip From", "")).strip()
        logic = str(row.get("Logic", "")).strip()
        skip_to = str(row.get("Skip To", "")).strip()

        if skip_to == "" or skip_from == "":
            continue

        # Always Skip = unconditional blank
        if str(row.get("Always Skip", "0")) == "1":
            rules.append({
                "Question": skip_to,
                "Check_Type": "Skip",
                "Condition": f"If 1=1 then {skip_to} should be blank"
            })
            continue

        if logic and logic.lower() != "nan":
            # Forward rule (if logic true -> skip_to blank)
            rules.append({
                "Question": skip_to,
                "Check_Type": "Skip",
                "Condition": f"If {logic} then {skip_to} should be blank"
            })
            # Reverse rule (if logic false -> skip_to answered)
            rules.append({
                "Question": skip_to,
                "Check_Type": "Skip",
                "Condition": f"If NOT({logic}) then {skip_to} should be answered"
            })
    return rules

# --- Step 2: Convert Constructed List File to Validation Rules ---
def convert_constructed_list(constructed_file):
    rules = []
    with open(constructed_file, "r", encoding="utf-8") as f:
        content = f.read()

    # Split by List Name
    blocks = re.split(r"List Name:", content)
    for block in blocks[1:]:
        lines = block.strip().splitlines()
        list_name = lines[0].strip()
        logic_lines = [l for l in lines if l.strip().startswith("if")]

        for logic in logic_lines:
            m = re.match(r"if\((.*?)\)\s*{(.*?)}", logic)
            if not m:
                continue
            condition, action = m.groups()
            # Extract ADD(PARENTLISTNAME(),val)
            add_match = re.search(r"ADD\(.*?,\s*(\d+)\)", action)
            if add_match:
                val = add_match.group(1)
                rules.append({
                    "Question": list_name,
                    "Check_Type": "Skip",
                    "Condition": f"If {condition} then {list_name}={val}"
                })
    return rules

# --- Step 3: Combine and Export ---
def convert_rules(skip_file, constructed_file, output_file="validation_rules.xlsx"):
    skip_rules = convert_skip_excel(skip_file)
    constructed_rules = convert_constructed_list(constructed_file)

    all_rules = skip_rules + constructed_rules
    rules_df = pd.DataFrame(all_rules)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        rules_df.to_excel(writer, index=False, sheet_name="Validation Rules")

    print(f"✅ Validation rules exported to {output_file}")

# --- Run Example ---
if __name__ == "__main__":
    convert_rules("skip_logic.xlsx", "Print Study.txt", "validation_rules.xlsx")
