import pandas as pd
import subprocess

# Execute pdftotext command to convert PDF to text
subprocess.run(["pdftotext", "BTECH1.pdf", "output.txt"])

# Initialize lists to store data
roll_numbers = []
names = []
sids = []

# Open the text file and read its contents line by line
with open('output.txt', 'r') as file:
    lines = file.readlines()

# Loop through the lines and extract roll number, name, and SID
for i in range(len(lines)):
    line = lines[i]
    if line.strip().isdigit() and len(line.strip()) == 11:  # Check if the line contains a valid roll number
        roll_numbers.append(line.strip())  # Add roll number to the list
        # Assume the next line contains the name
        if i + 1 < len(lines):
            name_line = lines[i + 1].strip()
            if not name_line.startswith('SID:'):
                names.append(name_line)  # Add name to the list
                if i + 2 < len(lines):
                    sid_line = lines[i + 2].strip()
                    if sid_line.startswith('SID:'):
                        sids.append(sid_line.split(':')[1].strip())  # Extract SID from the following line
                    else:
                        sids.append('')  # If SID line is missing
                else:
                    sids.append('')  # If SID line is missing
            else:
                names.append('')  # If name line is missing
                sids.append('')  # If SID line is missing

# Check if the lengths of the lists match
if len(roll_numbers) != len(names) or len(roll_numbers) != len(sids):
    print("Error: Lengths of roll_numbers, names, and sids lists do not match.")
    print(f"Roll Numbers: {len(roll_numbers)}, Names: {len(names)}, SIDs: {len(sids)}")
    exit()

# Create a DataFrame from the extracted data
data = {'Roll Number': roll_numbers, 'Name': names, 'SID': sids}
df = pd.DataFrame(data)

# Write the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
