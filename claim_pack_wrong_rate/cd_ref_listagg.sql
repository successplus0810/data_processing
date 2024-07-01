-- select item_idnt, clm_state, sum(clm_qty)  clm_qty, round(sum(clm_product)/sum(clm_qty),2)  clm_rate , sum(clm_product) clm_product
select item_idnt, clm_state, sum(clm_qty)  clm_qty, round(sum(clm_product)/NULLIFZERO(sum(clm_qty)),2)  clm_rate , sum(clm_product) clm_product
, listagg(distinct CLM_REF_NUM,', ') ref_num
from COLES.STI_WIP_CS.CLAIM_DETAILS_SUMMED_all
where item_idnt in ({})
and (clm_code like '%PS%' or clm_code like '%BS%' or clm_code like '%COLOTHER%')
and trim(PRMTN_COMP_IDNT) = '{}'
GROUP BY item_idnt, clm_state
ORDER BY item_idnt, clm_state