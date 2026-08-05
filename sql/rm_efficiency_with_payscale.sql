-- rm_efficiency_with_payscale.sql
-- Pulls raw per-RM performance data (Active POS, NOP, Gross Premium, Payin,
-- Payout, Retention) and full-joins it against monthly salary data, so
-- efficiency (retention generated per rupee of salary cost) can be
-- calculated downstream.

USE your_database;

------------------------ Pulls RAW data for efficiency along with Salary -------------------------

SELECT
    COALESCE(f.d_branch, s.Cost_Centre) AS Branch,
    COALESCE(s.Employee_Code, f.emp_id) AS Employee_Code,
    COALESCE(s.RM_code, f.rm_code) AS RM_code,
    s.Employee_Name,
    COALESCE(s.month_period, f.month_period) AS month_period,
    s.Total_salary,
    f.rm_name,
    f.Active_POS,
    f.NOP,
    f.gp,
    f.Payin,
    f.Payout,
    f.Retentions
FROM (
    SELECT
        Cost_Centre,
        Employee_Code,
        RM_code,
        Employee_Name,
        month_period,
        SUM(Overall_salary) AS Total_salary
    FROM salary_data
    WHERE dt_month_period BETWEEN '2025-04-01' AND '2025-08-31'
    GROUP BY Cost_Centre, Employee_Code, RM_code, Employee_Name, month_period
) s
FULL JOIN (
    SELECT
        d_branch,
        emp_id,
        rm_code,
        rm_name,
        month_period,
        COUNT(DISTINCT ref_pos_code) AS Active_POS,
        COUNT(nop) AS NOP,
        SUM(gross_premium) AS gp,
        SUM(payin_amount) AS Payin,
        SUM(payout_amount) AS Payout,
        SUM(payin_amount) - SUM(payout_amount) AS Retentions
    FROM final_all_data_dump
    WHERE start_date_a BETWEEN '2025-04-01' AND '2025-08-31'
        AND biz_type IN ('N', 'R')
    GROUP BY d_branch, emp_id, rm_code, rm_name, month_period
) f
ON s.Employee_Code = f.emp_id
    AND s.month_period = f.month_period;
