from collections import Counter
from dataclasses import dataclass, field
import sqlite3


@dataclass
class Audit:
    read: int = 0
    processed: int = 0
    excluded: Counter = field(default_factory=Counter)
    modes: Counter = field(default_factory=Counter)
    unknown: Counter = field(default_factory=Counter)
    issues: Counter = field(default_factory=Counter)
    duplicate_keys: int = 0
    samples: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    def controls(self, summary_total):
        return {
            'Registros leídos': self.read,
            'Registros RNC': self.processed,
            'Registros excluidos': sum(self.excluded.values()),
            'Total del resumen sin filtros': summary_total,
            'Diferencia leído - RNC - excluidos': self.read - self.processed - sum(self.excluded.values()),
            'Diferencia resumen - RNC': summary_total - self.processed,
            'Diferencia modalidades - RNC': sum(self.modes.values()) - self.processed,
            'Repeticiones originales de banco + AFIPOS (conservadas)': self.duplicate_keys,
            'Pertenencias no reconocidas (registros)': sum(self.unknown.values()),
            **{mode: count for mode, count in self.modes.items()},
        }

    def validate(self, summary_total):
        if self.read != self.processed + sum(self.excluded.values()):
            raise ValueError('No cuadra leído = RNC + excluidos.')
        if self.processed != summary_total or self.processed != sum(self.modes.values()):
            raise ValueError('El resumen o las modalidades no cuadran con RNC.')


class DuplicateTracker:
    """Disk-backed exact key tracking across chunks, no deletion of source rows."""
    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.execute('CREATE TABLE seen (bank TEXT, afipos TEXT, PRIMARY KEY (bank, afipos)) WITHOUT ROWID')

    def add(self, keys):
        keys = [(b, a) for b, a in keys if a]
        before = self.db.total_changes
        self.db.executemany('INSERT OR IGNORE INTO seen VALUES (?, ?)', keys)
        self.db.commit()
        return len(keys) - (self.db.total_changes - before)

    def close(self):
        self.db.close()
