select * 
from COLES.STI_WIP_CS.CLAIM_DETAILS_SUMMED_all
where item_idnt in ({})
and (clm_code like '%PS%' or clm_code like '%BS%' or clm_code like '%COLOTHER%')
and trim(PRMTN_COMP_IDNT) = '{}'
order by clm_end, item_idnt