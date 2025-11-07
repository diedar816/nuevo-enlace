--Script para la consulta de datos en Orfeo
select x.radi_nume_radi as NUMERO_RADICADO,TO_CHAR(X.radi_fech_radi,'DD/MM/YYYY') AS FECHA_ASIGNACION , X.COD_DEPENDENCIA,X.nombre_dependencia as nombre_dependencia,
(select trim(rn.nro_noticia) from radicado_notcrim rn where rn.radi_nume_radi = x.radi_nume_radi order by fecha_vinc desc FETCH FIRST 1 ROWS only) as nro_noticia_orfeo
from (
   SELECT    ra.radi_nume_radi as radi_nume_radi,ra.radi_fech_radi as radi_fech_radi, ra.radi_depe_radi AS COD_DEPENDENCIA,
           (select d.depe_nomb from dependencia d where d.depe_codi = ra.radi_depe_radi ) as nombre_dependencia
   FROM    public.radicado ra
   where    1 = 1
   ) x
where    1 = 1
and        X.radi_fech_radi >= to_date('2024/04/01 11:59:24','YYYY/MM/DD HH24:MI:SS') and X.radi_fech_radi < to_date('2025/07/01 11:59:59','YYYY/MM/DD HH24:MI:SS')
and    X.COD_DEPENDENCIA in (119,10,14,18,19,2,20,205,21,23,24,25,28,280,3,315,316,318,37,395,4,432,521,555,584,598,6,611,71,717,739,742,747,764,8,815,817,887,9,924)
order by X.radi_fech_radi
;
--Script para consultar los datos de SPOA, 
SELECT  C.CASO_USUARIO_CREACION, 
    (SELECT
		p.pers_primer_nombre || ' ' || p.pers_seg_nombre || ' ' || p.pers_primer_apellido || ' ' || p.pers_seg_apellido AS Nombre_Receptor_Denuncia
	FROM
		personas p,
		casos c_1,
		funcionarios f
	LEFT JOIN entidades et ON
		(f.FUNC_ENTI_ID = et.enti_id )
	LEFT JOIN direcciones_fiscalias df ON
		(f.FUNC_DIFI_ID = df.difi_id)
	LEFT JOIN seccionales_all se ON
		(f.FUNC_SECC_ID = se.secc_id)
	LEFT JOIN unidades_all un ON
		(f.FUNC_UNID_ID = un.unid_id)
	WHERE
		c.Caso_Usuario_Creacion = f.func_id_usuario
		AND c_1.caso_id = C.CASO_ID
		AND f.func_pers_id = p.pers_id) AS Nombre_Receptor_Denuncia,
    C.CASO_ID, C.CASO_NOTICIA,  
    TO_CHAR(C.CASO_FECHA_CREACION,'DD/MM/YYYY') AS CASO_FECHA_CREACION, 
    C.CASO_DELI_ID,
    (SELECT del.DELI_DESCRIPCION FROM DELITOS DEL WHERE DEL.DELI_ID = C.CASO_DELI_ID ) AS NOMBRE_DELITO,
    (SELECT
		se.secc_descripcion Seccional_Receptor
	FROM
		personas p,
		casos c_2,
		funcionarios f
	LEFT JOIN entidades et ON
		(f.FUNC_ENTI_ID = et.enti_id )
	LEFT JOIN direcciones_fiscalias df ON
		(f.FUNC_DIFI_ID = df.difi_id)
	LEFT JOIN seccionales_all se ON
		(f.FUNC_SECC_ID = se.secc_id)
	LEFT JOIN unidades_all un ON
		(f.FUNC_UNID_ID = un.unid_id)
	WHERE
		c.Caso_Usuario_Creacion = f.func_id_usuario
		AND c_2.caso_id = C.CASO_ID
		AND f.func_pers_id = p.pers_id) AS NOMBRE_SECCIONAL
FROM    CASOS C
WHERE   1 = 1
AND     C.CASO_FECHA_CREACION >= TO_DATE('01/04/2025 00:00:00', 'DD/MM/RRRR HH24:MI:SS') AND C.CASO_FECHA_CREACION < TO_DATE('30/06/2025 00:00:00', 'DD/MM/RRRR HH24:MI:SS');
--
