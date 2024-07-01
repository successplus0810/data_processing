select * 
from COLES.STI_WIP_CS.CLAIM_DETAILS_SUMMED_ALL
where item_idnt in ('{}')
and (clm_code like '%PS%' or clm_code like '%BS%' or clm_code like '%COLOTHER%')
and clm_code not ilike '%man%'
and clm_start between '{}'::date and '{}'::date
and clm_end between '{}'::date and '{}'::date
order by clm_end, item_idnt