SELECT ITEMID::integer item_idnt, FROM_DATE clm_start, TO_DATE clm_end, PROMOID PRMTN_COMP_IDNT, STATE clm_state, QUANTITY clm_qty, REBATE_RATE clm_rate, REBATE_AMOUNT clm_product, GST, REF_NUM clm_ref_num
FROM coles.STI_WIP_CE.CLAIM_DETAILS_SUM 
WHERE itemid IN ({})
and trim(promoid) = '{}'