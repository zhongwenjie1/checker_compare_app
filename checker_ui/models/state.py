import json
import os
from pathlib import Path
from typing import Dict, Optional


class AppState:
    def __init__(self):
        # 原有字段
        self.area_to_station: Dict[str, str] = {}
        self.station_cycle_times: Dict[str, int] = {}

        # P0：新增 岗位（zone/station）统一预设（不改变原有逻辑，只是增量能力）
        # 形如：{"电检": {"capacity": 2, "cycle_time": 118}, "电检准备": {"capacity": 1, "cycle_time": 118}}
        # capacity >= 1；cycle_time 为秒，>0 有效，None 表示不覆盖
        self.zone_configs: Dict[str, Dict[str, Optional[int]]] = {}

    # ---------------- 原有映射能力 ----------------

    def set_station_for_area(self, area_id: str, station_id: str) -> None:
        """把某个 area 映射到 station（岗位）。"""
        self.area_to_station[str(area_id)] = str(station_id)
        # 可在此处持久化：self.save_state()

    def remove_station_for_area(self, area_id: str) -> None:
        self.area_to_station.pop(str(area_id), None)
        # self.save_state()

    def set_cycle_time_for_station(self, station_id: str, seconds: int) -> None:
        """为某个岗位设置“显式节拍”（秒）。若设置了它，优先级高于 zone_configs 的统一节拍。"""
        try:
            seconds = int(seconds)
        except Exception:
            raise ValueError("seconds must be integer")
        self.station_cycle_times[str(station_id)] = seconds
        # self.save_state()

    def remove_cycle_time_for_station(self, station_id: str) -> None:
        self.station_cycle_times.pop(str(station_id), None)
        # self.save_state()

    def get_station_for_area(self, area_id: str) -> str:
        """返回 area 对应的 station；默认回退到自身，保证向后兼容。"""
        return self.area_to_station.get(str(area_id), str(area_id))

    def get_cycle_time(self, area_id: str) -> Optional[int]:
        """保留旧行为：仅读取“显式节拍映射”。（兼容旧逻辑）"""
        station = self.get_station_for_area(area_id)
        val = self.station_cycle_times.get(str(station))
        if val is None:
            return None
        return int(val)

    # ---------------- P0 新增：岗位统一预设（并行能力 + 统一节拍） ----------------

    def _sanitize_capacity(self, value) -> int:
        try:
            v = int(value)
            return v if v >= 1 else 1
        except Exception:
            return 1

    def _sanitize_cycle_time(self, value) -> Optional[int]:
        if value is None:
            return None
        try:
            v = int(value)
            if v > 0:
                return v
        except Exception:
            pass
        return None

    def set_zone_config(self, zone_id: str, *, capacity: Optional[int] = None, cycle_time: Optional[int] = None) -> None:
        """新增/更新某个岗位（zone/station）的统一预设。"""
        if not zone_id:
            return
        z = self.get_zone_config(zone_id)  # 取得清洗后的默认/现有值
        if capacity is not None:
            z["capacity"] = self._sanitize_capacity(capacity)
        if cycle_time is not None:
            z["cycle_time"] = self._sanitize_cycle_time(cycle_time)
        self.zone_configs[str(zone_id)] = z
        # self.save_state()

    def remove_zone_config(self, zone_id: str) -> None:
        self.zone_configs.pop(str(zone_id), None)
        # self.save_state()

    def get_zone_config(self, zone_id: str) -> Dict[str, Optional[int]]:
        """读取岗位统一预设，若不存在返回默认：capacity=1, cycle_time=None。"""
        raw = self.zone_configs.get(str(zone_id))
        if not isinstance(raw, dict):
            return {"capacity": 1, "cycle_time": None}
        cap = self._sanitize_capacity(raw.get("capacity", 1))
        ct = self._sanitize_cycle_time(raw.get("cycle_time", None))
        return {"capacity": cap, "cycle_time": ct}

    def get_effective_cycle_time(self, area_id: str) -> Optional[int]:
        """推荐新用法：优先取显式节拍，其次取岗位统一节拍；都无则 None。"""
        station = self.get_station_for_area(area_id)

        # 1) 显式节拍（旧映射）
        v = self.station_cycle_times.get(str(station))
        if v is not None:
            try:
                return int(v)
            except Exception:
                return None

        # 2) 岗位统一节拍（P0 预设）
        z = self.get_zone_config(station)
        return z.get("cycle_time", None)

    # ---------------- 持久化（保持对旧逻辑的向后兼容） ----------------

    def to_dict(self) -> dict:
        """
        序列化当前状态为 dict。
        若你有基类/旧实现，这里保留向后兼容的 try/except。
        """
        base = {}
        try:
            base = super().to_dict()  # 若无基类会抛异常，走 except
        except Exception:
            # 若项目里有其它持久化字段，请按你的实际合并
            base = getattr(self, "_persist_base", {}) or {}

        # 写入旧字段
        base["area_to_station"] = self.area_to_station
        base["station_cycle_times"] = self.station_cycle_times

        # 写入新增字段（已是可 JSON 序列化的纯 dict）
        base["zone_configs"] = self.zone_configs
        return base

    @classmethod
    def from_dict(cls, data: dict):
        """从 dict 反序列化。保持旧字段，再加载新增字段。"""
        inst = cls()  # 若你的 __init__ 需要参数，请按项目实际调整

        # 旧字段
        inst.area_to_station = data.get("area_to_station", {}) or {}
        inst.station_cycle_times = data.get("station_cycle_times", {}) or {}

        # 新字段：清洗一下再保存
        zc = data.get("zone_configs", {}) or {}
        cleaned: Dict[str, Dict[str, Optional[int]]] = {}
        if isinstance(zc, dict):
            for k, v in zc.items():
                if isinstance(v, dict):
                    cap = v.get("capacity", 1)
                    ct = v.get("cycle_time", None)
                    cleaned[str(k)] = {
                        "capacity": inst._sanitize_capacity(cap),
                        "cycle_time": inst._sanitize_cycle_time(ct),
                    }
        inst.zone_configs = cleaned
        return inst

    def save_state(self, path: Optional[str] = None) -> None:
        """示例持久化：写入 JSON 文件。项目已有保存逻辑时，把新增字段包含进去即可。"""
        if path is None:
            cfg_dir = Path.home() / ".checkerui"
            cfg_dir.mkdir(parents=True, exist_ok=True)
            path = str(cfg_dir / "config.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    def load_state(self, path: Optional[str] = None) -> None:
        """示例加载：把 zone_configs 一并读出。"""
        if path is None:
            cfg_dir = Path.home() / ".checkerui"
            path = str(cfg_dir / "config.json")
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # 旧字段
        self.area_to_station = data.get("area_to_station", {}) or {}
        self.station_cycle_times = data.get("station_cycle_times", {}) or {}

        # 新字段
        zc = data.get("zone_configs", {}) or {}
        cleaned: Dict[str, Dict[str, Optional[int]]] = {}
        if isinstance(zc, dict):
            for k, v in zc.items():
                if isinstance(v, dict):
                    cap = v.get("capacity", 1)
                    ct = v.get("cycle_time", None)
                    cleaned[str(k)] = {
                        "capacity": self._sanitize_capacity(cap),
                        "cycle_time": self._sanitize_cycle_time(ct),
                    }
        self.zone_configs = cleaned