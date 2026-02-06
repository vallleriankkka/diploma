from typing import List, Optional, Dict, Any
from datetime import datetime


class CommissionExpertDAO:
    """DAO для работы со связями комиссий и экспертов"""

    def __init__(self, db):
        self.db = db

    def update_expert_role(self, commission_id: int, expert_id: int, role_id: int) -> bool:
        """Обновление роли эксперта в комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE commission_experts 
                SET role_id = %s 
                WHERE commission_id = %s AND expert_id = %s
            """, (role_id, commission_id, expert_id))
            return cur.rowcount > 0

    def get_expert_commissions(self, expert_id: int) -> List[tuple]:
        """Получение комиссий, в которых состоит эксперт"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT ce.*, ec.name as commission_name, ec.description,
                       cr.name as role_name
                FROM commission_experts ce
                JOIN commissions ec ON ce.commission_id = ec.id
                JOIN commission_roles cr ON ce.role_id = cr.id
                WHERE ce.expert_id = %s
                ORDER BY ec.created_date DESC
            """, (expert_id,))
            return cur.fetchall()

    def is_expert_in_commission(self, commission_id: int, expert_id: int) -> bool:
        """Проверка, состоит ли эксперт в комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) FROM commission_experts 
                WHERE commission_id = %s AND expert_id = %s
            """, (commission_id, expert_id))
            result = cur.fetchone()
            return result[0] > 0 if result else False
