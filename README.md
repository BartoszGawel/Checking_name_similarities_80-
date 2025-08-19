# Checking_name_similarities_80-
Simple program to check if position name is similar to other from created dictionary in 80%.

Excel Job Title Cleaner & Standardizer
This Python script processes an Excel file containing job titles, and performs the following actions:

Counts Job Title Occurrences
It scans choosed column and counts how many times each job title appears.

Identifies Rare Titles
Job titles that appear less than twice are considered rare and are visually marked.

Highlights Rare Titles
Rare job titles are filled with a calm blue color to indicate they may need review.

Suggests Standardized Titles
For each rare title, the script checks if it is at least 80% similar to a more common title (based on string similarity). If so, it suggests the closest match in column C.

Applies Consistent Formatting
Suggested titles are formatted to:

Start with a capital letter (e.g. Manager, Sales Assistant)
Keep specific terms in uppercase (e.g. BHP)
Preserve brand names.

Make sure your job titles are in column B, starting from row 2.
Run the script. It will:

Highlight rare titles in blue
Suggest a standardized title in column C
The updated file will be saved in the same location with the same name (you can change this if needed).

Requirements:
Python 3.x
openpyxl
difflib
