SELECT * FROM TABLE (COLES.LIQUORLAND.CT_CHECK_GAP(STRTOK_TO_ARRAY('{}',','),DATE('{}','YYYY-MM-DD'),DATE('{}','YYYY-MM-DD'),STRTOK_TO_ARRAY('{}',','), STRTOK_TO_ARRAY('{}',','), '{}'))