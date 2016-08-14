PROCEDURE erz_reset
 GO TOP 
 SCAN 
  REPLACE erzst WITH 0
  _vfp.ActiveForm.refresh
 ENDSCAN 
 GO TOP 
RETURN 