-- rm_efficiency_monthly_pivot.sql
-- Month-wise multi-metric pivot per RM (SQL Server).
-- Pulls Active POS, NOP, Gross Premium, Payin, Payout and Retention for
-- each relationship manager, laid out as one column per metric-per-month
-- so it can be dropped straight into a BI tool or pivot report.

USE your_database;
GO

;WITH monthly AS (
    SELECT
        d_branch,
        rm_code,
        rm_name,
        month_period,
        COUNT(DISTINCT ref_pos_code)                AS Active_POS,
        COUNT(nop)                                  AS NOP,
        SUM(gross_premium)                          AS GP,
        SUM(payin_amount)                           AS PII,
        SUM(payout_amount)                          AS PO,
        SUM(payin_amount) - SUM(payout_amount)      AS RET
    FROM final_all_data_dump
    WHERE start_date_a BETWEEN '2025-04-01' AND '2025-10-31'
      AND biz_type IN ('N', 'R')
    GROUP BY d_branch, rm_code, rm_name, month_period
)
SELECT
    d_branch,
    rm_code,
    rm_name,

    -- Active POS
    SUM(CASE WHEN month_period = 'Apr-25' THEN Active_POS END) AS [ActivePOS_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN Active_POS END) AS [ActivePOS_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN Active_POS END) AS [ActivePOS_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN Active_POS END) AS [ActivePOS_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN Active_POS END) AS [ActivePOS_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN Active_POS END) AS [ActivePOS_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN Active_POS END) AS [ActivePOS_Oct-25],

    -- NOP (number of policies)
    SUM(CASE WHEN month_period = 'Apr-25' THEN NOP END) AS [NOP_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN NOP END) AS [NOP_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN NOP END) AS [NOP_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN NOP END) AS [NOP_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN NOP END) AS [NOP_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN NOP END) AS [NOP_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN NOP END) AS [NOP_Oct-25],

    -- Gross Premium
    SUM(CASE WHEN month_period = 'Apr-25' THEN GP END) AS [GP_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN GP END) AS [GP_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN GP END) AS [GP_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN GP END) AS [GP_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN GP END) AS [GP_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN GP END) AS [GP_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN GP END) AS [GP_Oct-25],

    -- Payin
    SUM(CASE WHEN month_period = 'Apr-25' THEN PII END) AS [PII_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN PII END) AS [PII_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN PII END) AS [PII_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN PII END) AS [PII_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN PII END) AS [PII_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN PII END) AS [PII_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN PII END) AS [PII_Oct-25],

    -- Payout
    SUM(CASE WHEN month_period = 'Apr-25' THEN PO END) AS [PO_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN PO END) AS [PO_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN PO END) AS [PO_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN PO END) AS [PO_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN PO END) AS [PO_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN PO END) AS [PO_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN PO END) AS [PO_Oct-25],

    -- Retention
    SUM(CASE WHEN month_period = 'Apr-25' THEN RET END) AS [RET_Apr-25],
    SUM(CASE WHEN month_period = 'May-25' THEN RET END) AS [RET_May-25],
    SUM(CASE WHEN month_period = 'Jun-25' THEN RET END) AS [RET_Jun-25],
    SUM(CASE WHEN month_period = 'Jul-25' THEN RET END) AS [RET_Jul-25],
    SUM(CASE WHEN month_period = 'Aug-25' THEN RET END) AS [RET_Aug-25],
    SUM(CASE WHEN month_period = 'Sep-25' THEN RET END) AS [RET_Sep-25],
    SUM(CASE WHEN month_period = 'Oct-25' THEN RET END) AS [RET_Oct-25]

FROM monthly
GROUP BY d_branch, rm_code, rm_name
ORDER BY d_branch;
