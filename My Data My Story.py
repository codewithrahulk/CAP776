import math
import openpyxl
from datetime import date

def read_file(file_name, sheet_name, start_date, end_date):
    wb = openpyxl.load_workbook(file_name, data_only=True)
    ws = wb[sheet_name]

    data = []

    for col in range(1, ws.max_column + 1):
        val = []
        for row in range(6, ws.max_row + 1):
            value = ws.cell(row, col).value
            if value is not None:
                val.append(value)
        data.append(val)
    return data

#average coding minutes per day
def calculate_tpi(data):
    tech_min = 0
    for val in data[4]:
        tech_min += val
    return round(tech_min / len(data[4]), 2)

#average study minutes per day
def calculate_study_time(data):
    study_min = 0
    for val in data[3]:
        study_min += val
    return round(study_min / len(data[3]), 2)

#average class minutes per day
def calculate_class_time(data):
    class_min = 0
    for val in data[5]:
        class_min += val
    return round(class_min / len(data[5]), 2)

#average study + class minutes per day
def calculate_aai(data):
    academic_min = 0
    for val in data[3]:
        academic_min += val
    for val in data[5]:
        academic_min += val
    return round(academic_min / len(data[5]), 3)


#average fitness minutes per day    
def calculate_PhAi(data):
    physical = 0
    for val in data[2]:
        physical += val
    return round(physical / len(data[2]), 2)

#average sleep minutes per day
def calculate_sri(data):
    sleep_recovery = 0
    for val in data[1]:
        sleep_recovery += val
    return round(sleep_recovery / len(data[1]), 3)

#avergae unaccounted minutes per day
def calculate_abi(data):
    unaccounted = 0
    for val in data[9]:
        unaccounted += val
    return round(unaccounted / len(data[9]), 2)

#average total tracked minutes per day
def calculate_tui(data):
    tracked = 0
    for val in data[8]:
        tracked += val
    return round(tracked / len(data[8]), 3)


feeling = {
    "Excellent" : 5,
    "Good" : 4,
    "Neutral" : 3,
    "Low" : 2,
    "Stressed" : 1
}

satisfaction = {
    "Very Satisfied" : 5,
    "Satisfied" : 4,
    "Neutral" : 3,
    "Unsatisfied" : 2,
    "Very Unsatisfied" : 1
}

energy = {
    "High" : 3,
    "Medium" : 2,
    "Low" : 1
}

#average feeling rating per day
def calculate_feeling_score(data):
    total = 0
    for val in data[10]:
        total += feeling[val]
    return round(total / len(data[10]), 2)

#average satisfaction rating per day
def calculate_satisfaction_score(data):
    total = 0
    for val in data[11]:
        total += satisfaction[val]
    return round(total / len(data[11]), 2)

#average energy rating per day
def calculate_energy_score(data):
    total = 0
    for val in data[12]:
        total += energy[val]
    return round(total / len(data[12]), 2)

#average feeling rating label per day
def get_feeling_level(score):
    if score == 5:
        return "Excellent"
    elif score >= 4:
        return "Good" 
    elif score >= 3:
        return "Neutral"
    elif score >= 2:
        return "Low"
    else:
        return "Stressed"

#average satisfaction rating label per day
def get_satisfaction_level(score):
    if score == 5:
        return "Very Satisfied"
    elif score >= 4:
        return "Satisfied" 
    elif score >= 3:
        return "Neutral"
    elif score >= 2:
        return "Unsatisfied"
    else:
        return "Very Unsatisfied"

#average energy rating label per day
def get_energy_level(score):
    if score == 3:
        return "High"
    elif score >= 2:
        return "Medium"
    else:
        return "Low"

#average of feeling, satisfaction and energy rating per day
def calculate_ei(data):
    experience = 0
    for val in data[10]:
        experience += feeling[val]
    for val in data[11]:
        experience += satisfaction[val]
    for val in data[12]:
        experience += energy[val]
    return round(experience / (3 * len(data[10])), 2)

def calculate_oti(data):
    sleep_recovery = 0
    for val in data[7]:
        sleep_recovery += val
    return round(sleep_recovery / len(data[7]), 3)

#percentage of valid days to expected days
def calculate_dci(valid_days, expected_days):
    return round((valid_days // expected_days) * 100, 2)

#weightage of all the indices to calculate overall performance assessment index
def calculate_pai(TPI, AAI, PhAi, SRI, ABI, TUI, EI, DCI):
    return round((0.15 * TPI) + (0.20 * AAI) + (0.15 * PhAi) + (0.20 * SRI) + (0.15 * TUI) + (0.10 * EI) + (0.05 * DCI), 3)

def calculate_mean(lst):
    return sum(lst) / len(lst) if lst else 0.0

def calculate_correlation(x, y):
    n = len(x)
    if n == 0 or len(y) != n:
        return 0.0

    mean_x = calculate_mean(x)
    mean_y = calculate_mean(y)

    numerator = sum((x[i] - mean_x) * (y[i] - mean_y) for i in range(n))
    sum_sq_x = sum((x[i] - mean_x) ** 2 for i in range(n))
    sum_sq_y = sum((y[i] - mean_y) ** 2 for i in range(n))

    denominator = math.sqrt(sum_sq_x * sum_sq_y)
    return numerator / denominator if denominator != 0 else 0.0

file_name = "12604978.xlsx"
sheet_name = "Daily Log"
start_date = date(2026, 8, 17)
end_date = date(2026, 9, 21)

data = read_file(file_name, sheet_name, start_date, end_date)

valid_days = len(data[0])
expected_days = (end_date - start_date).days + 1

TPI = calculate_tpi(data)
AAI = calculate_aai(data)
study_time = calculate_study_time(data)
class_time = calculate_class_time(data)

PhAi = calculate_PhAi(data)
SRI = calculate_sri(data)
ABI = calculate_abi(data)
TUI = calculate_tui(data)
EI = calculate_ei(data)

feeling_score = calculate_feeling_score(data)
satisfaction_score = calculate_satisfaction_score(data)
energy_score = calculate_energy_score(data)

feeling_label = get_feeling_level(feeling_score)
satisfaction_label = get_satisfaction_level(satisfaction_score)
energy_label = get_energy_level(energy_score)

energy_scores = [energy[val] for val in data[12]]
satisfaction_scores = [satisfaction[val] for val in data[11]]
feeling_scores = [feeling[val] for val in data[10]]

corr_sleep_energy = calculate_correlation(data[1], energy_scores)
corr_study_satisfaction = calculate_correlation(data[3], satisfaction_scores)
corr_coding_energy = calculate_correlation(data[4], energy_scores)
corr_fitness_energy = calculate_correlation(data[2], energy_scores)
corr_sleep_feeling = calculate_correlation(data[1], feeling_scores)

OTI = calculate_oti(data)
DCI = calculate_dci(valid_days, expected_days)
PAI = calculate_pai(TPI, AAI, PhAi, SRI, ABI, TUI, EI, DCI)

print("------ << MY DATA MY STORY >> ------")
print("\nExpected Days:", expected_days)
print("Valid Days:", valid_days)
print("\n------ ACTIVITY DATA SUMMARY ------")
print("1. Technical Proficiency Index (TPI):", TPI)
print("2. Academic Activity Index (AAI):", AAI)
print("  - Average Study Time:", study_time, "min/day")
print("  - Average Class Time:", class_time, "min/day")
print("3. Physical Activity (PhAi):", PhAi)
print("4. Sleep & Recovery Index (SRI):", SRI)
print("5. Activity Balance (ABI):", ABI)
print("6. Time Utilization Index (TUI):", TUI)
print("  - Other Activities Index (OTI):", OTI)
print("7. Experience Index (EI):", EI)
print(f"  - Average Feeling Score: {feeling_score} ({feeling_label})")
print(f"  - Average Satisfaction Score: {satisfaction_score} ({satisfaction_label})")
print(f"  - Average Energy Score: {energy_score} ({energy_label})")
print("8. Data Continuity Index (DCI):", DCI,"%")
print(" >> PERFORMANCE ASSESSMENT INDEX (PAI):", PAI)

print("\n------ CORRELATION ANALYSIS ------")
print(f"  Sleep vs Energy       : {corr_sleep_energy:.4f}")
print(f"  Study vs Satisfaction : {corr_study_satisfaction:.4f}")
print(f"  Coding vs Energy      : {corr_coding_energy:.4f}")
print(f"  Fitness vs Energy     : {corr_fitness_energy:.4f}")
print(f"  Sleep vs Feeling      : {corr_sleep_feeling:.4f}")