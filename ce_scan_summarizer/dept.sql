SELECT DISTINCT DEPT_IDNT 
FROM COLES_CLEAN.SUPERMARKET_MERCH.cml_prditmdm_daily 
WHERE item_idnt IN ({})
LIMIT 1
