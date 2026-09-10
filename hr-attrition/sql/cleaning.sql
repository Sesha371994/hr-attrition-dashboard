-- ============================================================
-- HR ATTRITION DATA CLEANING
-- ============================================================

-- 1. VALIDATION: measure the mess first
SELECT COUNT(*) AS total_rows FROM raw_hr;
SELECT COUNT(*) AS invalid_age FROM raw_hr WHERE age < 18 OR age > 70 OR age IS NULL;
SELECT COUNT(*) AS invalid_salary FROM raw_hr WHERE monthly_salary <= 0 OR monthly_salary IS NULL;
SELECT COUNT(*) AS invalid_satisfaction FROM raw_hr WHERE satisfaction_score < 1 OR satisfaction_score > 5 OR satisfaction_score IS NULL;
SELECT DISTINCT gender FROM raw_hr;
SELECT DISTINCT department FROM raw_hr;
SELECT DISTINCT attrited FROM raw_hr;

-- 2. CLEAN TABLE
DROP TABLE IF EXISTS clean_hr;
CREATE TABLE clean_hr AS
SELECT DISTINCT
  employee_id,
  TRIM(name) AS name,
  CASE
    WHEN LOWER(TRIM(gender)) IN ('male','m') THEN 'Male'
    WHEN LOWER(TRIM(gender)) IN ('female','f') THEN 'Female'
    ELSE 'Unspecified'
  END AS gender,
  CASE WHEN age BETWEEN 18 AND 70 THEN age ELSE NULL END AS age,
  UPPER(SUBSTR(TRIM(department),1,1)) || LOWER(SUBSTR(TRIM(department),2)) AS department,
  education,
  CASE WHEN monthly_salary > 0 THEN monthly_salary ELSE NULL END AS monthly_salary,
  join_date_raw,
  exit_date_raw,
  CASE WHEN satisfaction_score BETWEEN 1 AND 5 THEN satisfaction_score ELSE NULL END AS satisfaction_score,
  CASE WHEN LOWER(TRIM(attrited)) = 'yes' THEN 'Yes' ELSE 'No' END AS attrited
FROM (
  SELECT employee_id, name, gender, age, department, education, monthly_salary,
         join_date AS join_date_raw, exit_date AS exit_date_raw,
         satisfaction_score, attrited
  FROM raw_hr
);

-- 3. Validation after cleaning (should show clean counts)
SELECT COUNT(*) AS clean_row_count FROM clean_hr;
SELECT DISTINCT gender FROM clean_hr;
SELECT DISTINCT department FROM clean_hr;
