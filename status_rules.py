from config import MOVABLE_STATUSES, STATUS_PROGRESS, NEAR_FINISH_STATUSES


def get_status_progress(status):
    sl = str(status).lower()
    for key, val in STATUS_PROGRESS.items():
        if key in sl:
            return val
    return 0


def is_status(status_lower, status_list):
    for s in status_list:
        if s in status_lower:
            return True
    return False


def status_effective_weight(status):
    sl = str(status).lower()
    if 'cancelado' in sl or 'cancelada' in sl:
        return 0.35
    if any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada']):
        return 0.05
    if any(k in sl for k in NEAR_FINISH_STATUSES):
        return 0.65
    if get_status_progress(sl) >= 1:
        return 1.25
    if is_status(sl, MOVABLE_STATUSES):
        return 1.05
    return 0.95


def status_completion_credit(status):
    sl = str(status).lower()
    if any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada']):
        return 1.0
    if 'cancelado' in sl or 'cancelada' in sl:
        return 0.25
    return 0.0


def is_completed_or_auditable(status):
    sl = str(status).lower()
    return any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada'])


def should_alert_low_progress(tech_name, completed_count, peer_counts):
    if tech_name == 'SIN_ASIGNAR':
        return False
    if completed_count > 1:
        return False
    valid_peers = [c for c in peer_counts if c is not None]
    if len(valid_peers) < 2:
        return False
    peers_with_2_or_more = sum(1 for c in valid_peers if c >= 2)
    peers_with_3_or_more = sum(1 for c in valid_peers if c >= 3)
    peer_avg = sum(valid_peers) / len(valid_peers) if valid_peers else 0
    return peers_with_2_or_more >= 2 and (peers_with_3_or_more >= 1 or peer_avg >= 2.0)
