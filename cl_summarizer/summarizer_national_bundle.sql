SELECT DATE1, ITEMIDSKU, ITEMNAME, BRANDID, UNITOFMEASURE UOM_QTY, STATE, ITEMPRICE, RRP_EXC, TOTALLINESALE,CAST(TOTAL_SALE_QTY AS int) TOTAL_SALE_QTY, AVG_SALE_PRICE, CAST(NON_PROMO_QTY AS int) NON_PROMO_QTY, NON_PROMO_AVG_SALE_PRICE, TOTALLINESALE_PROMO, CAST(PROMO_QTY AS int) PROMO_QTY, PROMO_AVG_SALE_PRICE
, {} AS SCAN FROM TABLE (COLES.LIQUORLAND.BUNDLE_CL_MULTIITEM( ARRAY_CONSTRUCT({}), '{}', '{}', '{}', {}, {}, {}+0.2::NUMBER(38,2))) 
--SCAN-----------------------------------------------------------------------ITEM-- START--END--BRAND-UOM--BQTY-BRRP------------
order by brandid, ITEMIDSKU , UOM_QTY , DATE1