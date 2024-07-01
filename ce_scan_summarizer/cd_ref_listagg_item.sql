select item_idnt, sum(clm_qty)  clm_qty, round(sum(clm_product)/sum(clm_qty),2)  clm_rate , sum(clm_product) clm_product
, listagg(distinct CLM_REF_NUM,', ') ref_num
from 
(SELECT ITEMID::integer item_idnt, FROM_DATE clm_start, TO_DATE clm_end, PROMOID PRMTN_COMP_IDNT, STATE clm_state, QUANTITY clm_qty, REBATE_RATE clm_rate, REBATE_AMOUNT clm_product, GST, REF_NUM clm_ref_num
FROM coles.STI_WIP_CE.CLAIM_DETAILS_SUM 
WHERE itemid in ('{}')
and FROM_DATE between '{}'::date and '{}'::date
and TO_DATE between '{}'::date and '{}'::date)
GROUP BY item_idnt
ORDER BY item_idnt