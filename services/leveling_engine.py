"""
services/leveling_engine.py
Motor central de nivelacion de appointments.
Entrada: lista de dicts (de Metabase o Excel).
Salida: dict JSON con resumen, alertas, sugerencias, carga por tecnico y franja.
Sin dependencias de openpyxl ni generacion de archivos Excel.
"""
import logging
from datetime import datetime, timedelta
from config import (
    MAX_IDEAL_LOAD, MIN_IDEAL_LOAD, MAX_ABSOLUTE_LOAD, MAX_ORDERS_PER_SLOT,
    MAX_DUPLICATED_SLOTS, MIN_IMBALANCE_TO_MOVE, ORDER_DURATION_HOURS,
    MAX_ORDER_DURATION_HOURS, MAX_ALLOWED_DISTANCE_KM, FRAGMENTATION_PENALTY,
    INTERZONE_DISTANCE_PENALTY, ZONE_ONLY_NO_COORDS_PENALTY,
    MAX_INTERZONE_ASSIGNMENTS_PER_TECH, ZONE_ADJACENCY, MOVABLE_STATUSES,
    ONSITE_ALERT_MINUTES, OVERLOAD_PER_SLOT, now_bogota,
)
from services.normalization import (
    normalize_order, haversine, get_centroid, is_same_unit, order_has_coords,
    parse_franja_hours, get_status_progress, status_effective_weight,
    status_completion_credit, norm_zone, is_movable, is_blocked,
    norm_status,
)

logger = logging.getLogger(__name__)

def _count_dup(fc): return sum(1 for c in fc.values() if c >= 2)

def _parse_dt(s):
    if not s: return None
    if hasattr(s, "replace"):
        try: return s.replace(tzinfo=None)
        except: pass
    for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%dT%H:%M:%S","%Y-%m-%d %H:%M","%d/%m/%Y %H:%M"):
        try: return datetime.strptime(str(s).strip()[:19],fmt)
        except: pass
    return None

def _tech_ref(tech,to,tl):
    for o in to.get(tech,[]):
        if 1<=o.get("progress",0)<6 and o.get("lat") and o.get("lon"): return(o["lat"],o["lon"])
    up=[(parse_franja_hours(o.get("franja",""))[0] or 99,o) for o in to.get(tech,[]) if o.get("movible") and o.get("lat") and o.get("lon")]
    if up: up.sort(key=lambda x:x[0]); od=up[0][1]; return(od["lat"],od["lon"])
    locs=tl.get(tech,[])
    return get_centroid(locs) if locs else(0.0,0.0)

def _dist(order,tech,to,tl):
    if not order.get("lat") or not order.get("lon"): return None
    ref=_tech_ref(tech,to,tl)
    return haversine(order["lat"],order["lon"],ref[0],ref[1]) if ref!=(0.0,0.0) else None

def _can_add(tech,franja,tf,to):
    cur=tf.get(tech,{}).get(franja,0)
    if cur>=MAX_ORDERS_PER_SLOT: return False
    if cur>=1 and _count_dup(tf.get(tech,{}))>=MAX_DUPLICATED_SLOTS: return False
    fs,_=parse_franja_hours(franja)
    if fs and fs>=14.5:
        t=sum(c for f,c in tf.get(tech,{}).items() if (parse_franja_hours(f)[0] or 0)>=14.5)
        if t>=2: return False
    return True

def _build_idx(orders):
    to={};tf={};ts={};tl={};tz={};zt={}
    for o in orders:
        t=o["tecnico"];f=o["franja"];z=o["zona"];sz=o["subzona"]
        to.setdefault(t,[]).append(o)
        tf.setdefault(t,{}); tf[t][f]=tf[t].get(f,0)+1
        ts.setdefault(t,set()).add(sz)
        if o.get("lat") and o.get("lon"): tl.setdefault(t,[]).append((o["lat"],o["lon"]))
        tz.setdefault(t,{}); tz[t][z]=tz[t].get(z,0)+1
        zt.setdefault(z,set()).add(t)
    tmz={t:max(zones.items(),key=lambda x:x[1])[0] for t,zones in tz.items()}
    tt={t:len(ol) for t,ol in to.items()}
    tp={t:sum(1 for o in ol if o["movible"]) for t,ol in to.items()}
    tel={t:sum(o["effective_weight"] for o in ol) for t,ol in to.items()}
    tc={t:sum(o["completion_credit"] for o in ol) for t,ol in to.items()}
    return {"tech_orders":to,"tech_franja":tf,"tech_subzones":ts,"tech_locs":tl,
            "tech_main_zone":tmz,"zone_techs":{z:list(ts2) for z,ts2 in zt.items()},
            "tech_total":tt,"tech_pending":tp,"tech_eff_load":tel,"tech_credit":tc}

def _alerts(orders,idx,now_dt):
    alerts=[]; now_h=now_dt.hour+now_dt.minute/60.0
    for o in orders:
        if norm_status(o["estado"])=="en sitio":
            upd=_parse_dt(o.get("updated_at",""))
            if upd:
                mins=(now_dt.replace(tzinfo=None)-upd).total_seconds()/60
                if mins>ONSITE_ALERT_MINUTES:
                    alerts.append({"tipo":"EN_SITIO_PROLONGADO","severidad":"critica" if mins>ONSITE_ALERT_MINUTES*2 else "alta",
                        "orden":o["id"],"tecnico":o["tecnico"],"franja":o["franja"],"zona":o["zona"],
                        "detalle":f"Orden {o['id']} en sitio {int(mins)} min"})
    for tech,fm in idx["tech_franja"].items():
        if tech=="SIN_ASIGNAR": continue
        for franja,count in fm.items():
            if count>=OVERLOAD_PER_SLOT:
                alerts.append({"tipo":"SOBRECARGA_FRANJA","severidad":"alta" if count>OVERLOAD_PER_SLOT else "media",
                    "tecnico":tech,"franja":franja,"count":count,"detalle":f"{tech} {count} ords en {franja}"})
    for tech,total in idx["tech_total"].items():
        if tech=="SIN_ASIGNAR": continue
        if total>MAX_ABSOLUTE_LOAD:
            alerts.append({"tipo":"SOBRECARGA_TOTAL","severidad":"alta","tecnico":tech,"total":total,"detalle":f"{tech} tiene {total} ords (max:{MAX_ABSOLUTE_LOAD})"})
        elif total>MAX_IDEAL_LOAD:
            alerts.append({"tipo":"SOBRECARGA_TOTAL","severidad":"media","tecnico":tech,"total":total,"detalle":f"{tech} tiene {total} ords (techo:{MAX_IDEAL_LOAD})"})
    for tech,total in idx["tech_total"].items():
        if tech=="SIN_ASIGNAR": continue
        if total<MIN_IDEAL_LOAD and idx["tech_pending"].get(tech,0)>0:
            alerts.append({"tipo":"CARGA_BAJA","severidad":"media","tecnico":tech,"total":total,"detalle":f"{tech} solo {total} ords (min:{MIN_IDEAL_LOAD})"})
    for o in orders:
        if o["movible"]:
            if o["tecnico"]=="SIN_ASIGNAR":
                alerts.append({"tipo":"SIN_TECNICO","severidad":"media","orden":o["id"],"franja":o["franja"],"zona":o["zona"],"detalle":f"Orden {o['id']} sin tecnico"})
            elif o["franja"]=="Sin Franja":
                alerts.append({"tipo":"SIN_FRANJA","severidad":"media","orden":o["id"],"tecnico":o["tecnico"],"zona":o["zona"],"detalle":f"Orden {o['id']} sin franja"})
    for o in orders:
        if o["movible"] and o["franja"]!="Sin Franja":
            _,fe=parse_franja_hours(o["franja"])
            if fe is not None and fe<now_h-0.5:
                alerts.append({"tipo":"FRANJA_VENCIDA","severidad":"alta","orden":o["id"],"tecnico":o["tecnico"],"franja":o["franja"],"detalle":f"Orden {o['id']} franja vencida"})
    return alerts

def _score(order,donor,receiver,idx):
    to=idx["tech_orders"];tf=idx["tech_franja"];tl=idx["tech_locs"];tt=idx["tech_total"]
    dt=tt.get(donor,0); rt=tt.get(receiver,0)
    if rt>=MAX_ABSOLUTE_LOAD: return -9999.0
    sob=(rt>=MAX_IDEAL_LOAD)
    gap=dt-rt
    if gap<MIN_IMBALANCE_TO_MOVE and donor not in ("SIN_ASIGNAR",None): return -9999.0
    score=gap*800
    if sob: score-=1500
    if rt<MIN_IDEAL_LOAD: score+=(MIN_IDEAL_LOAD-rt)*600
    if dt>MAX_IDEAL_LOAD: score+=(dt-MAX_IDEAL_LOAD)*500
    if not _can_add(receiver,order["franja"],tf,to): return -9999.0
    score+=300
    dr=_dist(order,receiver,to,tl)
    dd=_dist(order,donor,to,tl) if donor not in ("SIN_ASIGNAR",None) else None
    if dr and dr>MAX_ALLOWED_DISTANCE_KM: score-=INTERZONE_DISTANCE_PENALTY
    if dr is None: score-=200
    if dd and dr: score+=(dd-dr)*150
    rs=len(idx["tech_subzones"].get(receiver,set()))
    if order["subzona"] not in idx["tech_subzones"].get(receiver,set()) and rs>=3: score-=FRAGMENTATION_PENALTY*0.3
    return score

def _suggestions(orders,idx):
    sugs=[]; movables=[o for o in orders if o["movible"]]
    techs=[t for t in idx["tech_orders"] if t!="SIN_ASIGNAR"]
    to=idx["tech_orders"];tf=idx["tech_franja"];tt=idx["tech_total"]
    ts=idx["tech_subzones"];tl=idx["tech_locs"];tmz=idx["tech_main_zone"]
    iz={}
    for order in movables:
        donor=order["tecnico"]; dt=tt.get(donor,0)
        if donor not in ("SIN_ASIGNAR",None) and dt<=MAX_IDEAL_LOAD:
            if not any(tt.get(t,0)<MIN_IDEAL_LOAD for t in techs if t!=donor): continue
        best_s=-9999.0; best_r=None; best_risk="bajo"
        for r in techs:
            if r==donor: continue
            if tt.get(r,0)>=MAX_ABSOLUTE_LOAD: continue
            s=_score(order,donor,r,idx)
            if s>best_s: best_s=s; best_r=r
        if best_r is None or best_s<0: continue
        dv=tt.get(donor,0); rv=tt.get(best_r,0)
        oz=order["zona"]; rz=tmz.get(best_r,"SIN_ZONA")
        interz=(rz!=oz and oz not in ZONE_ADJACENCY.get(rz,[rz]))
        if interz:
            best_risk="alto"; iz[best_r]=iz.get(best_r,0)+1
            if iz[best_r]>MAX_INTERZONE_ASSIGNMENTS_PER_TECH: continue
        dr=_dist(order,best_r,to,tl); dd=_dist(order,donor,to,tl) if donor!="SIN_ASIGNAR" else None
        sob=(rv>=MAX_IDEAL_LOAD)
        if donor=="SIN_ASIGNAR": motivo=f"Sin tecnico -> {best_r}({rv} ords)"
        elif dv>MAX_IDEAL_LOAD: motivo=f"Sobrecarga: {donor}({dv})->{best_r}({rv})"
        elif rv<MIN_IDEAL_LOAD: motivo=f"Deficit: {best_r} tiene {rv} ords (min {MIN_IDEAL_LOAD})"
        else: motivo=f"Balance: {donor}({dv})->{best_r}({rv})"
        if sob: motivo+=f" AVISO: {best_r} quedara con {rv+1} ords"
        ben=[]
        if dv>MAX_IDEAL_LOAD: ben.append(f"Descarga {donor}")
        if rv<MIN_IDEAL_LOAD: ben.append(f"Completa cuota {best_r}")
        if dd and dr and dd>dr: ben.append(f"Ahorro ~{dd-dr:.1f}km")
        if not ben: ben.append("Mejora balance")
        sugs.append({"orden":order["id"],"tecnico_actual":donor,"tecnico_sugerido":best_r,
            "franja_actual":order["franja"],"franja_sugerida":order["franja"],
            "tipo":order.get("tipo",""),"estado":order["estado"],"zona":oz,"motivo":motivo,
            "riesgo":best_risk,"beneficio":" / ".join(ben),"score":round(best_s,1),
            "interzona":interz,"aviso_sobrecarga":sob,"total_receptor":rv,"total_donante":dv,
            "dist_receptor_km":round(dr,2) if dr else None})
    sugs.sort(key=lambda x:x["score"],reverse=True)
    return sugs[:50]

def run_leveling(raw_orders):
    now_dt=now_bogota(); now_h=now_dt.hour+now_dt.minute/60.0
    if not raw_orders: return _empty("Sin datos. Configura Metabase o sube un Excel.")
    orders=[normalize_order(o) for o in raw_orders]
    idx=_build_idx(orders)
    movibles=[o for o in orders if o["movible"]]
    bloqueadas=[o for o in orders if not o["movible"]]
    alerts=_alerts(orders,idx,now_dt)
    sugs=_suggestions(orders,idx)
    cpt=[]
    for t,ol in sorted(idx["tech_orders"].items()):
        tot=len(ol); mov=sum(1 for o in ol if o["movible"])
        act=sum(1 for o in ol if 1<=o.get("progress",0)<6)
        fin=sum(1 for o in ol if o.get("progress",0)>=6)
        cpt.append({"tecnico":t,"zona":idx["tech_main_zone"].get(t,"SIN_ZONA"),"total":tot,
            "movibles":mov,"bloqueadas":tot-mov,"activas":act,"finalizadas":fin,
            "sobrecarga":tot>MAX_IDEAL_LOAD,"franjas":idx["tech_franja"].get(t,{}),
            "subzonas":list(idx["tech_subzones"].get(t,set()))})
    from config import FRANJAS
    fd={f:{"total":0,"movibles":0,"bloqueadas":0,"tecnicos":set()} for f in FRANJAS}
    fd["Sin Franja"]={"total":0,"movibles":0,"bloqueadas":0,"tecnicos":set()}
    for o in orders:
        f=o["franja"]; fd.setdefault(f,{"total":0,"movibles":0,"bloqueadas":0,"tecnicos":set()})
        fd[f]["total"]+=1
        if o["movible"]: fd[f]["movibles"]+=1
        else: fd[f]["bloqueadas"]+=1
        if o["tecnico"]!="SIN_ASIGNAR": fd[f]["tecnicos"].add(o["tecnico"])
    cpf=[{"franja":f,"total":d["total"],"movibles":d["movibles"],"bloqueadas":d["bloqueadas"],
          "tecnicos":len(d["tecnicos"]),"sobrecarga":d["total"]>MAX_ORDERS_PER_SLOT*max(len(d["tecnicos"]),1)}
         for f,d in fd.items() if d["total"]>0]
    def en(o): return {k:o.get(k,"") for k in ["id","tecnico","estado","estado_clase","franja","tipo","zona","subzona","direccion","movible","progress","updated_at","lat","lon"]}
    al_crit=sum(1 for a in alerts if a.get("severidad") in ("critica","alta"))
    ts_count=len([t for t in idx["tech_orders"] if t!="SIN_ASIGNAR"])
    tsob=sum(1 for c in cpt if c["sobrecarga"] and c["tecnico"]!="SIN_ASIGNAR")
    tcap=sum(1 for t,tot in idx["tech_total"].items() if t!="SIN_ASIGNAR" and tot<MAX_IDEAL_LOAD)
    tdef=sum(1 for t,tot in idx["tech_total"].items() if t!="SIN_ASIGNAR" and tot<MIN_IDEAL_LOAD)
    return {"generado_en":now_dt.strftime("%Y-%m-%d %H:%M:%S"),
        "resumen":{"total_ordenes":len(orders),"movibles":len(movibles),"bloqueadas":len(bloqueadas),
            "alertas":len(alerts),"alertas_criticas":al_crit,"sugerencias":len(sugs),
            "tecnicos_total":ts_count,"tecnicos_sobrecargados":tsob,"tecnicos_con_capacidad":tcap,
            "tecnicos_deficitarios":tdef,
            "sin_tecnico":sum(1 for o in movibles if o["tecnico"]=="SIN_ASIGNAR"),
            "sin_franja":sum(1 for o in movibles if o["franja"]=="Sin Franja"),
            "objetivo_por_tecnico":f"{MIN_IDEAL_LOAD}-{MAX_IDEAL_LOAD} ordenes"},
        "carga_por_tecnico":cpt,"carga_por_franja":cpf,
        "ordenes_movibles":[en(o) for o in movibles],"ordenes_bloqueadas":[en(o) for o in bloqueadas],
        "alertas":alerts,"sugerencias":sugs}

def _empty(msg):
    return {"generado_en":now_bogota().strftime("%Y-%m-%d %H:%M:%S"),"mensaje":msg,
        "resumen":{"total_ordenes":0,"movibles":0,"bloqueadas":0,"alertas":0,"alertas_criticas":0,
            "sugerencias":0,"tecnicos_total":0,"tecnicos_sobrecargados":0,"tecnicos_con_capacidad":0,
            "tecnicos_deficitarios":0,"sin_tecnico":0,"sin_franja":0,
            "objetivo_por_tecnico":f"{MIN_IDEAL_LOAD}-{MAX_IDEAL_LOAD} ordenes"},
        "carga_por_tecnico":[],"carga_por_franja":[],"ordenes_movibles":[],"ordenes_bloqueadas":[],
        "alertas":[],"sugerencias":[]}
