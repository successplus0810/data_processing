SELECT ITEMIDSKU, MODE(ITEMNAME) itemname, MODE(ITEMGROUP) itemgroup
	FROM COLES_CLEAN.LIQUORLAND_MERCH.ALL_DIMSALEITEM
	WHERE itemidsku IN ('{}')
	and lower(itemgroup) in ('beer','non','rtd','spirits','wine')
	GROUP BY ITEMIDSKU