select x.radi_nume_radi as NUMERO_RADICADO,X.radi_fech_radi as FECHA_ASIGNACION, X.COD_DEPENDENCIA,X.nombre_dependencia as nombre_dependencia,
(select trim(rn.nro_noticia) from radicado_notcrim rn where rn.radi_nume_radi = x.radi_nume_radi) as nro_noticia_orfeo
from (
	SELECT 	ra.radi_nume_radi as radi_nume_radi,ra.radi_fech_radi as radi_fech_radi, ra.radi_depe_radi AS COD_DEPENDENCIA,
			(select d.depe_nomb from dependencia d where d.depe_codi = ra.radi_depe_radi) as nombre_dependencia
	FROM 	public.radicado ra
	where 	1 = 1 
	) x
where 	1 = 1
and	  	X.radi_fech_radi >= to_date('2024/06/30 11:10:24','YYYY/MM/DD HH24:MI:SS')
order by X.radi_fech_radi
;
