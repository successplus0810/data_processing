SELECT STATE, BRANDID, DESCRIPTION, STARTDATE, ENDDATE, REBATESCANLINE, REBATENO, ITEMID, MULTIPLIER_NUM  UOM_QTY, CRITERIA, REBATEREFID, REFTABLEID, REFRECID, REBATEPOSTED, CLAIMREFID, CLAIMTABLEID, PRODUCTCATEGORY, REVERSALCLAIMREFID, CLM_PER_UNIT, CLM_QTY/MULTIPLIER_NUM CLM_QTY, CLM_VAL, CD_TYPE, REBATE_ENTITLEMENT_NUM, ENTITLEMENT_AMT, CLM_WET, REBATESCANCLASS, REVERSED, VENDOR_NUMBER, VENDOR_NAME 
FROM COLES.LIQUORLAND.CD_FULL 
-- LEFT JOIN COLES.LIQUORLAND.UNITOFMEASURE_TRANSLATE B 
-- ON A.UNITOFMEASURE = B.UNITOFMEASURE
WHERE STARTDATE BETWEEN '{}' AND '{}' 
-- AND ENDDATE BETWEEN '{}' AND '{}' 
AND ITEMID = '{}' AND BRANDID = '{}' AND UPPER(CD_TYPE) = 'SCANDOLLAR'  AND MULTIPLIER_NUM = '{}' AND REBATE_ENTITLEMENT_NUM = '{}' AND CLAIMREFID IS NOT NULL AND LENGTH(TRIM(CLAIMREFID)) >0 AND (LENGTH(TRIM(REVERSALCLAIMREFID)) < 1 OR REVERSALCLAIMREFID IS NULL) 
-- format(----------startdate--enddate---------------startdate--enddate----------itemid------------,brandid,-------------------------------------------------------uom------------------------------,scan)