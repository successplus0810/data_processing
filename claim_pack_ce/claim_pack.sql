SELECT SLS_START, SLS_END, CLM_START, CLM_END, PRMTN_COMP_IDNT, PRMTN_COMP_NAME, SKU_ID, ITEM_DESC,state, DEPT_IDNT, CML_COST_GST_RATE_PCT, VENDOR_NUM, SUPPLIER, SUPP_DESC, CLM_RATE, PRM_PRICE , PRM_QTY, CLM_QTY, F_ACTUAL_SELL_PRICE_AMT_MOD, F_NORM_SELL_PRICE_AMT, VAR_QTY, ELI_ITEM, ELI,
case when  ELI_EXCLUDE is null then 0 else ELI_EXCLUDE end ELI_EXCLUDE
 , PAF_LOCATION, EMAIL, check_prgx , IDENTIFIED_AMT, GAP_PROF_PRGX
FROM COLES.STI_WIP_CE.CT_PROMOQTY_FINAL_3
WHERE SUPP_DESC LIKE '%ARNOTT%'
AND clm_end between '2022-6-30' and '2023-5-30'
AND SUPPLIER IS NOT NULL