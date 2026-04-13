from dataclasses import dataclass, field
from typing import Optional, Dict, List
import pandas as pd

@dataclass
class AppState:
    std_df: Optional[pd.DataFrame] = None
    sys_df: Optional[pd.DataFrame] = None
    result_df: Optional[pd.DataFrame] = None
    visible_cols: List[str] = field(default_factory=list)
    meta: Dict[str, str] = field(default_factory=dict)