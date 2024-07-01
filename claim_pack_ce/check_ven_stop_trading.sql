select distinct RMS_NUM from
(SELECT RMS_NUM, max(PROCESS_DTE) AS max_process_date
FROM COLES.STI_WIP_CS.AP_DEAL_CMNT_ALL
GROUP BY RMS_NUM
having max_process_date < '2021-5-1')