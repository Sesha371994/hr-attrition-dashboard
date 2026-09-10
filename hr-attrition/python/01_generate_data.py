"""
Generate realistic messy HR dataset — employee records with attrition,
salary, department, tenure. Intentional data issues included.
"""
import csv
import random
from datetime import datetime, timedelta

random.seed(7)

departments = ["Sales", "sales", "SALES", "Engineering", "engineering",
               "Marketing", "HR", "Finance", "Operations", "operations"]
dept_clean = ["Sales", "Engineering", "Marketing", "HR", "Finance", "Operations"]

first_names = ["Arjun", "Divya", "Karthik", "Meera", "Suresh", "Anitha",
               "Ramesh", "Priya", "Vijay", "Kavya", "Naveen", "Sneha",
               "Ganesh", "Deepa", "Manoj", "Latha"]
last_names = ["Kumar", "Raj", "Sharma", "Nair", "Iyer", "Reddy", "Pillai"]

genders = ["Male", "Female", "male", "female", "M", "F"]
education = ["Bachelors", "Masters", "PhD", "bachelors", "MASTERS", None]

rows = []
for i in range(1, 251):
    fname = random.choice(first_names)
    lname = random.choice(last_names)
    dept = random.choice(departments)
    join_date = datetime(2018, 1, 1) + timedelta(days=random.randint(0, 2400))
    date_fmt = random.choice(["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y"])

    is_attrited = random.random() < 0.22  # ~22% attrition
    if is_attrited:
        exit_date = join_date + timedelta(days=random.randint(90, 1800))
        exit_date_str = exit_date.strftime(date_fmt) if exit_date < datetime(2025, 9, 1) else ""
    else:
        exit_date_str = ""

    salary = random.choice([
        random.randint(25000, 120000),
        None,  # missing salary
        -5000,  # bad data
    ])

    age = random.choice([random.randint(22, 58), None, 15, 99])  # some invalid ages

    satisfaction = random.choice([1, 2, 3, 4, 5, None, 0])  # 0 is invalid (scale is 1-5)

    rows.append({
        "employee_id": f"E{i:04d}",
        "name": f"  {fname} {lname}" if random.random() < 0.08 else f"{fname} {lname}",
        "gender": random.choice(genders),
        "age": age,
        "department": dept,
        "education": random.choice(education),
        "monthly_salary": salary,
        "join_date": join_date.strftime(date_fmt),
        "exit_date": exit_date_str,
        "satisfaction_score": satisfaction,
        "attrited": "Yes" if is_attrited else random.choice(["No", "no", "NO"]),
    })

# inject duplicate rows
for _ in range(12):
    rows.append(random.choice(rows).copy())

with open("raw_data/hr_employees.csv", "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=rows[0].keys())
    writer.writeheader()
    writer.writerows(rows)

print(f"Generated {len(rows)} employee records -> raw_data/hr_employees.csv")
