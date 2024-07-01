%pip install openpyxl
import openpyxl
import pandas as pd

# Load Excel files
results_file_path = 'my_result.xlsx'
courses_file_path = 'courses.xlsx'

# Load results data from the 'results' file
df_first_year = pd.read_excel(results_file_path, sheet_name='first', header=0)
df_second_year = pd.read_excel(results_file_path, sheet_name='second', header=0)
df_third_year = pd.read_excel(results_file_path, sheet_name='third', header=0)

# Load course data from the 'courses' file
grade_point_map_df = pd.read_excel(courses_file_path, sheet_name='grade point map', header=0)
path_1_courses_df = pd.read_excel(courses_file_path, sheet_name='path 1 courses', header=0)
compulsory_path_1_df = pd.read_excel(courses_file_path, sheet_name='compulsory path 1', header=0)
compulsory_path_2_df = pd.read_excel(courses_file_path, sheet_name='compulsory path 2', header=0)
get_gpa_eligible_courses_df = pd.read_excel(courses_file_path, sheet_name='get gpa eligibile courses', header=0)
valid_grades_df = pd.read_excel(courses_file_path, sheet_name='valid grades', header=0)

# Convert dataframes to dictionaries or lists 
grade_point_map = dict(zip(grade_point_map_df['Grade'], grade_point_map_df['Grade Point']))
path_1_courses = path_1_courses_df['Path 1 Courses'].str.strip().str.upper().tolist()
compulsory_path_1 = compulsory_path_1_df['Compulsory Path 1'].str.strip().str.upper().tolist()
compulsory_path_2 = compulsory_path_2_df['Compulsory Path 2'].str.strip().str.upper().tolist()
get_gpa_eligible_courses = get_gpa_eligible_courses_df['GPA Ineligible Courses'].str.strip().str.upper().tolist()
valid_grades = set(valid_grades_df['Valid Grades'].tolist())

# Validate course codes length
def validate_course_codes(course_list, course_list_name):
    invalid_courses = [course for course in course_list if len(course) != 10]
    return invalid_courses

invalid_courses_dict = {
    "Path 1 Courses": validate_course_codes(path_1_courses, "Path 1 Courses"),
    "Compulsory Path 1": validate_course_codes(compulsory_path_1, "Compulsory Path 1"),
    "Compulsory Path 2": validate_course_codes(compulsory_path_2, "Compulsory Path 2"),
    "Get GPA Eligible Courses": validate_course_codes(get_gpa_eligible_courses, "Get GPA Eligible Courses")
}

# Extract credits from course code
def extract_credits(course_code):
    # Check if course code ends with a digit
    if course_code[-1].isdigit():
        return int(course_code[-1]) 
    else:
        return 0 

# Determine student's path based on courses taken
def determine_path(df):
    if any(course.strip().upper() in df['Course Code'].str.strip().str.upper().values for course in path_1_courses):
        return 1
    return 2

# Check compulsory courses are completed
def check_compulsory_courses(df, path):
    if path == 1:
        compulsory_courses = compulsory_path_1
    else:
        compulsory_courses = compulsory_path_2
    completed_courses = df['Course Code'].str.strip().str.upper().values
    incomplete_courses = [course for course in compulsory_courses if course not in completed_courses]
    return all(course in completed_courses for course in compulsory_courses), incomplete_courses

# Filter out non-GPA eligible courses and courses with NaN grades
def filter_courses(df):
    df_filtered = df[~df['Course Code'].str.strip().str.upper().isin(get_gpa_eligible_courses)]
    df_filtered = df_filtered[df_filtered['Grade'].notna()]
    return df_filtered.copy()

# Function to calculate GPA
def calculate_gpa(df):
    df = df.copy()
    df['Credits'] = df['Course Code'].apply(extract_credits)
    df['Grade Points'] = df['Grade'].map(grade_point_map)
    df = df.sort_values(by=['Course Code', 'Grade Points', 'AcYear'], ascending=[True, False, True])
    df = df.drop_duplicates(subset=['Course Code'], keep='first')
    df['Credit Points'] = df['Credits'] * df['Grade Points']
    total_credits = df['Credits'].sum()
    total_credit_points = df['Credit Points'].sum()
    gpa = total_credit_points / total_credits if total_credits != 0 else 0
    return round(gpa, 2), total_credits

def check_grade_validity(df):
    error_grades = df[~df['Grade'].isin(valid_grades) & ~df['Grade'].isna()]['Grade'].unique()

    # Check if there are any grades marked as "Withheld"
    if 'Withheld' in df['Grade'].values:
        return True, error_grades, [], []
    else:
        previous_results = []  
        pending_results = []   

        grouped = df.groupby(['Course Code', 'Course Name'])

        for (course_code, course_name), group_df in grouped:
            if len(group_df) == 3:
                group_df['Grade Points'] = group_df['Grade'].map(grade_point_map)

                # Check if all grade points are different
                if len(group_df['Grade Points'].unique()) == 2:
                    # If all grade points are different, skip this group
                    continue

            if len(group_df) > 2:
                # Consider these entries as previous results
                group_df['Grade Points'] = group_df['Grade'].map(grade_point_map)
                sorted_group_df = group_df.sort_values(by=['Grade Points', 'AcYear'], ascending=[False, True])
                previous_grade = sorted_group_df.iloc[0]['Grade']
                ac_year = sorted_group_df.iloc[0]['AcYear']
                previous_results.append([course_code, course_name, previous_grade, ac_year])
            else:
                if pd.isna(group_df['Grade'].iloc[0]) or group_df['Grade'].iloc[0] == "":
                    pending_results.append([course_code, course_name, group_df['AcYear'].iloc[0]])
        return False, error_grades, previous_results, pending_results
        
# Combine data from all years
all_years_df = pd.concat([df_first_year, df_second_year, df_third_year])

# Drop duplicates based on specified columns
all_years_df.drop_duplicates(subset=['Course Code', 'Course Name', 'AcYear', 'Attempt', 'Grade'], inplace=True)

all_years_df['Course Code'] = all_years_df['Course Code'].str.strip().str.upper()

path = determine_path(all_years_df)

# Check all compulsory courses are completed
compulsory_completed, incomplete_compulsory_courses = check_compulsory_courses(all_years_df, path)

# Filter out non-GPA eligible courses and courses with NaN grades
df_first_year_filtered = filter_courses(df_first_year)
df_second_year_filtered = filter_courses(df_second_year)
df_third_year_filtered = filter_courses(df_third_year)

# Calculate GPA for each year
gpa_first_year, credits_first_year = calculate_gpa(df_first_year_filtered)
gpa_second_year, credits_second_year = calculate_gpa(df_second_year_filtered)
gpa_third_year, credits_third_year = calculate_gpa(df_third_year_filtered)

# Calculate overall GPA
total_credits = credits_first_year + credits_second_year + credits_third_year
overall_gpa = (
    (gpa_first_year * credits_first_year) +
    (gpa_second_year * credits_second_year) +
    (gpa_third_year * credits_third_year)
) / total_credits if total_credits != 0 else 0
overall_gpa = round(overall_gpa, 2)

# Check grade validity and suspension status
suspended, error_grades, previous_results, pending_results = check_grade_validity(all_years_df)

# Check if ACLT and CMSK courses are completed
aclt_completed = any(course in all_years_df['Course Code'].values for course in ['ACLT 11012', 'ACLT 12022', 'ACLT 21032'])
cmsk_completed = any(course in all_years_df['Course Code'].values for course in ['CMSK 14012', 'CMSK 14022', 'CMSK 14032', 'CMSK 14042'])

# Check credit requirements
first_second_year_credits = credits_first_year + credits_second_year
third_year_credits = credits_third_year
total_credits_all_years = total_credits
elec_credits = all_years_df[all_years_df['Course Code'].str.startswith('ELEC') & all_years_df['Grade'].notna()]['Course Code'].apply(extract_credits).sum()
phys_credits = all_years_df[all_years_df['Course Code'].str.startswith('PHYS') & all_years_df['Grade'].notna()]['Course Code'].apply(extract_credits).sum()
credits_c_or_better = all_years_df[all_years_df['Grade'].map(grade_point_map) >= 2.0]['Course Code'].apply(extract_credits).sum()
credits_d_or_better = all_years_df[all_years_df['Grade'].map(grade_point_map) >= 1.0]['Course Code'].apply(extract_credits).sum()

# Determine degree eligibility
degree_eligibility = (
    compulsory_completed and aclt_completed and cmsk_completed and 
    (credits_first_year > 30) and (credits_second_year > 30) and 
    (first_second_year_credits > 60) and (third_year_credits > 30) and
    (total_credits_all_years > 90) and (overall_gpa >= 2.00) and
    (elec_credits >= 24) and (phys_credits >= 24) and ((elec_credits + phys_credits) >= 48) and
    (credits_c_or_better >= 72) and (credits_d_or_better >= 90)
)

# Determine class
degree_class = ""
credits_a_or_better = all_years_df[all_years_df['Grade'].map(grade_point_map) >= 4.0]['Course Code'].apply(extract_credits).sum()
credits_b_or_better = all_years_df[all_years_df['Grade'].map(grade_point_map) >= 3.0]['Course Code'].apply(extract_credits).sum()
if degree_eligibility:
    if (credits_c_or_better >= 90) and (credits_a_or_better >= (total_credits_all_years / 2)) and (overall_gpa >= 3.70):
        degree_class = "First Class"
    elif (credits_c_or_better >= 80) and (credits_d_or_better >= total_credits_all_years) and (credits_b_or_better >= (total_credits_all_years / 2)) and (overall_gpa >= 3.30):
        degree_class = "Second Class (Upper Division)"
    elif (credits_c_or_better >= 80) and (credits_d_or_better >= total_credits_all_years) and (credits_b_or_better >= (total_credits_all_years / 2)) and (overall_gpa >= 3.00):
        degree_class = "Second Class (Lower Division)"
    else:
        degree_class = "No Class Obtained"
else:
    degree_class = "Eligibility not met"

# Output data
reasons = []

# Check degree eligibility and reasons
if not compulsory_completed:
    reasons.append("Not all compulsory courses are completed.")
if not aclt_completed:
    reasons.append("ACLT course requirement is not fulfilled.")
if not cmsk_completed:
    reasons.append("CMSK course requirement is not fulfilled.")
if credits_first_year <= 30:
    reasons.append("Total credits for the first year is not greater than 30.")
if credits_second_year <= 30:
    reasons.append("Total credits for the second year is not greater than 30.")
if first_second_year_credits <= 60:
    reasons.append("Total credits for the first and second years is not greater than 60.")
if third_year_credits <= 30:
    reasons.append("Total credits for the third year is not greater than 30.")
if total_credits_all_years <= 90:
    reasons.append("Total credits for all three years is not greater than 90.")
if overall_gpa < 2.00:
    reasons.append("Overall GPA is not greater than 2.00.")
if elec_credits < 24:
    reasons.append("Total ELEC credits are not greater than or equal to 24.")
if phys_credits < 24:
    reasons.append("Total PHYS credits are not greater than or equal to 24.")
if (elec_credits + phys_credits) < 48:
    reasons.append("Total ELEC and PHYS credits are not greater than or equal to 48.")
if credits_c_or_better < 72:
    reasons.append("Less than 72 credits with grade C or better.")
if credits_d_or_better < 90:
    reasons.append("Less than 90 credits with grade D or better.")

# Determine degree eligibility
degree_eligibility = len(reasons) == 0

# Change suspension status to "Suspended" if there are grades marked as "Withheld"
if suspended:
    suspension_status = "Suspended"
else:
    suspension_status = "Not Suspended"

# Output data
gpa_data = {
    'Year': ['First Year', 'Second Year', 'Third Year', 'Overall'],
    'GPA': [gpa_first_year, gpa_second_year, gpa_third_year, overall_gpa],
    'Degree Eligibility': [degree_eligibility, '', '', ''],
    'Reason': [', '.join(reasons) if reasons else '', '', '', ''],
    'Class': [degree_class, '', '', ''],
    'Suspension Status': [suspension_status, '', '', '']
}

# Write data to Excel
with pd.ExcelWriter(results_file_path, engine='openpyxl', mode='a') as writer:
    gpa_df = pd.DataFrame(gpa_data)
    gpa_df.to_excel(writer, sheet_name='GPA and Eligibility', index=False)

    if invalid_courses_dict:
        invalid_courses_df = pd.DataFrame(invalid_courses_dict)
        invalid_courses_df.to_excel(writer, sheet_name='Invalid Courses', index=False)

    if incomplete_compulsory_courses:
        incomplete_compulsory_courses_df = pd.DataFrame({'Incomplete Compulsory Courses': incomplete_compulsory_courses})
        incomplete_compulsory_courses_df.to_excel(writer, sheet_name='Incomplete Compulsory Courses', index=False)

    if error_grades.size > 0:
        error_grades_df = pd.DataFrame({'Error Grades': error_grades})
        error_grades_df.to_excel(writer, sheet_name='Error Grades', index=False)

    if previous_results:
        previous_results_df = pd.DataFrame(previous_results, columns=['Course Code', 'Course Name', 'Previous Grade', 'AcYear'])
        previous_results_df.to_excel(writer, sheet_name='Previous Results', index=False)

    if pending_results:
        pending_results_df = pd.DataFrame(pending_results, columns=['Course Code', 'Course Name', 'AcYear'])
        pending_results_df.to_excel(writer, sheet_name='Pending Results', index=False)