# -*- coding: utf-8 -*-
"""
组合票排程 + 导出
- Zone（阻塞区域）：同一 zone_id 的连续步骤视为同一区域，容量=可同时处于区域的车辆数。
- gate_zone（闸门）+ gate_buffer（缓冲）：对上游某些步骤设 gate_zone=某区域，
  表示该步骤“开工前/放行时”需要考虑该区域前的“闸门缓冲”是否已满。
  gate_buffer=缓冲允许在“闸门 → 区域入口”链路上同时存在的在制车数量（默认=2）。
  这样可实现：2 号车能继续到电检1，但 3 号车需等 2 号车进入电检2 后才开始电检准备。
"""

from __future__ import annotations
import math
import heapq
from typing import List, Dict, Any, Tuple
import pandas as pd
from checker_ui.models.state import AppState


# ---------------- Excel 引擎选择 ---------------- #

def _choose_engine():
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl"
        except Exception:
            return None

# ---------------- 在调度前应用“区域→岗位”覆盖 ---------------- #
def _apply_area_station_overrides(defs: List[Dict[str, Any]], state: AppState) -> List[Dict[str, Any]]:
    """
    将每一行定义中的 zone_id / gate_zone_id 用岗位ID（station_id）覆盖；
    并且把岗位节拍（station_cycle_times）写入到 cycle_time 字段。
    仅改变调度用的标识，不改变实际 dur（作业时长）。
    """
    a2s: Dict[str, str] = getattr(state, "area_to_station", {}) or {}
    sct: Dict[str, float] = getattr(state, "station_cycle_times", {}) or {}

    new_defs: List[Dict[str, Any]] = []
    for d in defs:
        nd = dict(d)  # 复制，避免原数据被修改
        zid = str(nd.get("zone_id", "") or "").strip()
        if zid:
            sid = str(a2s.get(zid, zid)).strip()
            nd["zone_id"] = sid
            # 若该岗位配置了节拍，覆盖写入 cycle_time（仅影响区域入口的节拍限流）
            ct = sct.get(sid, None)
            if ct not in (None, ""):
                try:
                    nd["cycle_time"] = float(ct)
                except Exception:
                    pass

        gz = str(nd.get("gate_zone_id", "") or "").strip()
        if gz:
            nd["gate_zone_id"] = str(a2s.get(gz, gz)).strip()

        new_defs.append(nd)
    return new_defs


 # -------------- 自动推断 规则（闸门/缓冲/容量） -------------- #
def _auto_infer_rules(defs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    输入：UI 收集的步骤 defs（可能没有 gate / gate_buffer / zone_capacity）
    输出：填充了默认闸门、缓冲和区域容量的 defs（原地修改并返回）
    规则：
      - 若某行有 zone_id 且未给 zone_capacity，则用 capacity 作为 zone_capacity（默认1）
      - 从下往上扫描，维护“下游最近的 zone_id”。对没有 zone_id 的行：
          若存在下游 zone_id = Z，则自动设置 gate_zone_id = Z，gate_buffer = 2
    """
    n = len(defs)
    if n == 0:
        return defs

    # 1) 区域容量默认：zone_capacity <- capacity
    for rec in defs:
        zid = str(rec.get("zone_id") or "").strip()
        if zid:
            if "zone_capacity" not in rec or rec.get("zone_capacity") in (None, "", 0, "0"):
                try:
                    cap = max(1, int(float(rec.get("capacity", 1))))
                except Exception:
                    cap = 1
                rec["zone_capacity"] = cap

    # 2) 闸门/缓冲自动：自下而上找“下游最近的区域”
    next_zone_ahead: str | None = None
    for i in range(n - 1, -1, -1):
        rec = defs[i]
        zid = str(rec.get("zone_id") or "").strip()
        if zid:
            next_zone_ahead = zid
            # 本行属于某个区域，不自动加 gate
            continue

        # 本行不在任何区域里：若下游存在区域，则加 gate 到那个区域
        if next_zone_ahead:
            if not rec.get("gate_zone_id"):
                rec["gate_zone_id"] = next_zone_ahead
            # 默认缓冲：若未显式给，设为 2
            if not rec.get("gate_buffer"):
                rec["gate_buffer"] = 2

    return defs

# ---------------- 新增：选主节拍 ---------------- #
def _pick_primary_cycle(step_defs: List[Dict[str, Any]]) -> float:
    """
    从 defs 中选择“主节拍”（秒）：
      1) 先按 zone_id 聚合（同一区域多行：取最大值，更保守）；
      2) 在所有区域节拍中：若有众数，取众数；否则取最大值；
      3) 若完全没有，返回 0（不画竖线）。
    """
    from collections import Counter

    zone_ct: Dict[str, float] = {}
    for d in step_defs:
        zid = str(d.get("zone_id", "") or "").strip()
        ct = d.get("cycle_time", None)
        try:
            ct = float(ct) if ct not in (None, "") else None
        except Exception:
            ct = None
        if not zid or ct is None or ct <= 0:
            continue
        zone_ct[zid] = max(zone_ct.get(zid, 0.0), float(ct))

    if not zone_ct:
        return 0.0

    vals = list(zone_ct.values())
    # 统计众数：考虑浮点误差，四舍五入到 1 位小数后计数
    rounded = [round(v, 1) for v in vals]
    cnt = Counter(rounded)
    most, freq = cnt.most_common(1)[0]
    if len(cnt) == 1 or (len(cnt) > 1 and freq > sorted(cnt.values())[-2]):
        return float(most)
    return float(max(vals))


# ---------------- 解析步骤与 Zone ---------------- #
def _normalize_defs(step_defs: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Any]], Dict[str, int], Dict[str, float]]:
    """
    返回 (steps, zones, gate_buffers, zone_cycles)

    steps: 每步 {seq, display, group, duration, zone_id, gate_zone_id, cycle_time}
    zones: {zid: {"capacity": int, "first_seq": int, "last_seq": int}}
    gate_buffers: {zid: gate_buffer_int}  # 若未显式提供，默认=2
    zone_cycles: {zid: float}  # 区域节拍（秒）；仅在该区域入口步生效
    """
    steps: List[Dict[str, Any]] = []
    # 收集 gate_buffer（按 zone 聚合，取出现的最大值；默认=2）
    gate_buffers: Dict[str, int] = {}
    # 收集每个区域的节拍（取该区域出现的最大值，更保守）
    zone_cycles: Dict[str, float] = {}

    for d in step_defs:
        display = str(d.get("display", "")).strip()
        group = str(d.get("group", "")).strip() or display
        durations = list(d.get("durations", []))
        if not display or not durations:
            continue
        dur = float(durations[0])

        # 读取节拍字段（不覆盖实际作业时长）
        cycle_time_val = d.get("cycle_time", None)
        ct: float | None = None
        if cycle_time_val not in (None, ""):
            try:
                ct = float(cycle_time_val)
                if not (ct > 0):
                    ct = None
            except Exception:
                ct = None

        zone_id = str(d.get("zone_id", "") or "").strip()
        gate_zone_id = str(d.get("gate_zone_id", "") or "").strip()

        # 聚合节拍到 zone（仅按区域入口使用，不改变条形长度）
        if zone_id and ct is not None:
            if zone_id in zone_cycles:
                zone_cycles[zone_id] = max(zone_cycles[zone_id], ct)
            else:
                zone_cycles[zone_id] = ct

        # 聚合 gate_buffer
        if gate_zone_id:
            gb = d.get("gate_buffer", None)
            if gb is None:
                gb = 2  # 默认缓冲=2，符合“2号车可走，3号车等”的现场规则
            try:
                gb = max(1, int(float(gb)))
            except Exception:
                gb = 2
            if gate_zone_id in gate_buffers:
                gate_buffers[gate_zone_id] = max(gate_buffers[gate_zone_id], gb)
            else:
                gate_buffers[gate_zone_id] = gb

        steps.append({
            "seq": int(d.get("seq", len(steps) + 1)),
            "display": display,
            "group": group,
            "duration": dur,
            "zone_id": zone_id,
            "gate_zone_id": gate_zone_id,
            "cycle_time": ct,  # 区域节拍（秒），仅在区域入口限流使用
        })

    steps.sort(key=lambda x: x["seq"])
    if not steps:
        raise ValueError("没有有效的步骤定义")

    # 汇总 Zone：确定起止步骤序号 + 容量
    zones: Dict[str, Dict[str, Any]] = {}
    for s in steps:
        zid = s.get("zone_id", "")
        if not zid:
            continue
        z = zones.setdefault(zid, {"capacity": 1, "first_seq": s["seq"], "last_seq": s["seq"]})
        z["first_seq"] = min(z["first_seq"], s["seq"])
        z["last_seq"] = max(z["last_seq"], s["seq"])

    # 从原始定义里补全 zone 容量
    for d in step_defs:
        zid = str(d.get("zone_id", "") or "").strip()
        if not zid or zid not in zones:
            continue
        zcap = d.get("zone_capacity", None)
        if zcap is not None:
            try:
                zones[zid]["capacity"] = max(int(zones[zid]["capacity"]), int(zcap))
            except Exception:
                pass

    # gate_buffer 若某个 gate_zone 没出现在任何行里，忽略；出现但没显式给值时已默认为 2
    return steps, zones, gate_buffers, zone_cycles


# ---------------- 调度（含 Zone + gate_buffer） ---------------- #
def schedule(step_defs: List[Dict[str, Any]], cars: int) -> Tuple[List[Dict[str, Any]], float]:
    """
    返回：
      rows: 每车-每步记录：
        {car, step_seq, step_display, group, dur, start, svc_finish, depart, block_wait}
      max_time: 全局最后 depart
    """
    steps, zones, gate_buffers, zone_cycles = _normalize_defs(step_defs)
    m = len(steps)

    # 工位释放时刻（考虑阻塞传递）
    server_free = [0.0 for _ in range(m)]

    # Zone 名额堆：zid -> [free_time, ...]（长度=capacity）
    zone_heaps: Dict[str, List[float]] = {}
    for zid, zinfo in zones.items():
        cap = int(zinfo.get("capacity", 1)) or 1
        zone_heaps[zid] = [0.0 for _ in range(cap)]
        heapq.heapify(zone_heaps[zid])

    # 闸门缓冲：对每个 gate_zone 维护“尚未进入该 zone 的车辆的预计进入时刻”最小堆
    # pre_heap[z] 中的元素是“已经通过闸门但尚未进入 zone 的车辆的 ‘zone 入口开始时间’ ”
    pre_heap: Dict[str, List[float]] = {}

    # 记录每个区域上一次车辆进入该区域入口的时刻（用于节拍约束）
    last_entry_time: Dict[str, float] = {}

    def is_zone_entry(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["first_seq"]

    def is_zone_exit(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["last_seq"]

    rows: List[Dict[str, Any]] = []
    max_time = 0.0

    for car in range(1, cars + 1):
        prev_depart = 0.0
        # 记录该车是否经过某个 gate_zone（用于之后把它的“进入 zone 的时刻”加入 pre_heap）
        car_gate_zones: set[str] = set()

        for j, st in enumerate(steps):
            # ---- 计算本步开始时间：受上一步 depart、本步服务器空闲约束 ----
            start = max(server_free[j], prev_depart)

            # ---- 闸门缓冲约束（在 start 阶段判断）：允许“闸门→区域入口”链路上最多 gate_buffer 辆 ----
            gz = st.get("gate_zone_id", "")
            if gz:
                car_gate_zones.add(gz)
                # 取得该 gate_zone 的缓冲与堆
                gb = max(1, int(gate_buffers.get(gz, 2)))
                heap = pre_heap.setdefault(gz, [])

                # 移除所有“进入 zone 的时刻 <= start”的条目（这些车在 start 时刻已进入 zone，不再占用缓冲）
                while heap and heap[0] <= start:
                    heapq.heappop(heap)

                # 若缓冲已满（heap 大小 >= gb），则把 start 推迟到“最早一辆进入 zone 的时刻”
                # 推迟后再次清理（可能一次就够，也可能要多次）
                while len(heap) >= gb:
                    start = max(start, heap[0])
                    while heap and heap[0] <= start:
                        heapq.heappop(heap)

            # ---- 区域入口节拍约束（仅在区域入口步生效，不改变服务时长） ----
            if is_zone_entry(j):
                zid = st.get("zone_id", "")
                ct = float(zone_cycles.get(zid, 0.0))
                if ct > 0 and zid in last_entry_time:
                    start = max(start, last_entry_time[zid] + ct)

            # ---- 服务结束 ----
            svc_finish = start + float(st["duration"])

            # ---- depart 受“下步可接收（服务器释放 + zone 容量）”约束 ----
            if j < m - 1:
                next_ready = server_free[j + 1]

                # 若“下步”是某 Zone 的入口，还得等该 Zone 出现名额
                if is_zone_entry(j + 1):
                    nzid = steps[j + 1]["zone_id"]
                    nheap = zone_heaps[nzid]
                    next_ready = max(next_ready, nheap[0] if nheap else 0.0)

                # 注：闸门缓冲只在 start 阶段处理，不再额外卡 depart

                depart = max(svc_finish, next_ready)
            else:
                depart = svc_finish

            block_wait = max(0.0, depart - svc_finish)

            rows.append({
                "car": car,
                "step_seq": st["seq"],
                "step_display": st["display"],
                "group": st["group"],
                "dur": float(st["duration"]),
                "start": start,
                "svc_finish": svc_finish,
                "depart": depart,
                "block_wait": block_wait,
            })

            # ---- Zone 名额占用/释放 ----
            # 进入 Zone：仅在“Zone 入口步骤”发生，占用一个名额
            if is_zone_entry(j):
                zid = st["zone_id"]
                # 如果该车之前通过过指向该 zid 的某个闸门，则把它“进入 zone 的时刻（=本步 start=上一步 depart）”加入 pre_heap
                if zid in car_gate_zones:
                    heap = pre_heap.setdefault(zid, [])
                    heapq.heappush(heap, start)  # 之后其他车的“闸门开始”会受这个时间点约束

                # 记录该区域本次进站时刻（用于下一辆车的节拍限制）
                if float(zone_cycles.get(zid, 0.0)) > 0:
                    last_entry_time[zid] = start

                heap = zone_heaps[zid]
                if heap:
                    heapq.heappop(heap)  # 占用一个 zone 名额

            # 离开 Zone：仅在“Zone 最后一步”释放一个名额（名额释放时刻=本步 depart）
            if is_zone_exit(j):
                zid = st["zone_id"]
                heap = zone_heaps[zid]
                heapq.heappush(heap, depart)

            # ---- 更新状态，进入下一步 ----
            server_free[j] = depart
            prev_depart = depart
            max_time = max(max_time, depart)

    return rows, max_time


# ---------------- 等待统计 ---------------- #
def _build_car_slices(rows: List[Dict[str, Any]]) -> Dict[int, List[Dict[str, Any]]]:
    by_car: Dict[int, List[Dict[str, Any]]] = {}
    for r in rows:
        by_car.setdefault(r["car"], []).append(r)
    for k in by_car:
        by_car[k].sort(key=lambda x: x["step_seq"])
    return by_car


def _compute_entry_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """入站等待：车 i 第一步 start - 车 i-1 第一步 depart（<0 计 0）"""
    wait_map: Dict[int, float] = {}
    prev_first_depart = 0.0
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        first_start = steps[0]["start"]
        wait_map[car] = max(0.0, first_start - prev_first_depart)
        prev_first_depart = steps[0]["depart"]
    return wait_map


def _compute_total_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """总等待 = 入站等待 + Σ block_wait（所有步）"""
    entry_map = _compute_entry_wait(by_car)
    total_map: Dict[int, float] = {}
    for car, steps in by_car.items():
        inter = sum(max(0.0, s.get("block_wait", 0.0)) for s in steps)
        total_map[car] = float(entry_map.get(car, 0.0) + inter)
    return total_map


# ---------------- 导出入口 ---------------- #
def schedule_and_export(defs: List[Dict[str, Any]],
                        cars: int,
                        grid_step: float,
                        wait_policy: str,   # "before"/"after" 仅影响是否绘入站等待条
                        project: str,
                        dst_path: str) -> None:
    grid_step = 1.0 if (not isinstance(grid_step, (int, float)) or grid_step <= 0) else float(grid_step)

    # 先自动推断 gate/缓冲/容量（让用户只填基础信息也能跑）
    defs = _auto_infer_rules(defs)

    # 在调度前，应用“区域→岗位”覆盖与岗位节拍
    try:
        _state = AppState()
        _state.load_state()  # 若本地没有保存会静默使用默认空映射
        defs = _apply_area_station_overrides(defs, _state)
    except Exception:
        # 任何异常都不影响原有流程（保持向后兼容）
        pass

    rows, max_finish = schedule(defs, cars)

    # ---- 新增：选主节拍（自动） ----
    primary_cycle = _pick_primary_cycle(defs)  # 0 表示不画竖线

    # ---- 收集用户自定义颜色 (display -> hex)
    step_color_map = {d.get("display"): d.get("color") for d in defs if d.get("color")}

    engine = _choose_engine()
    if engine is None:
        raise RuntimeError("未找到可用的 Excel 引擎，请安装 xlsxwriter 或 openpyxl")

    if engine == "xlsxwriter":
        _export_with_xlsxwriter(rows, max_finish, grid_step, wait_policy, project, dst_path, step_color_map, primary_cycle)
    else:
        _export_with_openpyxl(rows, max_finish, grid_step, wait_policy, project, dst_path, primary_cycle)


# ---------------- 样式与工具 ---------------- #
def _palette():
    group_colors = [
        "#4CAF50", "#2196F3", "#9C27B0", "#FF9800", "#009688",
        "#795548", "#3F51B5", "#E91E63", "#00BCD4", "#8BC34A",
    ]
    wait_color = "#FFC107"
    return group_colors, wait_color


def _fmt_num(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.1f}"


# ---------------- xlsxwriter 彩色导出 ---------------- #
def _export_with_xlsxwriter(rows: List[Dict[str, Any]], max_finish: float,
                             grid_step: float, wait_policy: str,
                             project: str, dst_path: str,
                             step_color_map: Dict[str, str],
                             primary_cycle: float) -> None:
    import xlsxwriter  # type: ignore

    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    n_cols_grid = max(1, int(math.ceil(max_finish / grid_step)))

    with pd.ExcelWriter(dst_path, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("作业组合票")
        writer.sheets["作业组合票"] = ws

        fmt_header = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#EEEEEE", "border": 1})
        fmt_text   = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        fmt_left   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1})
        fmt_wait   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1, "bg_color": "#FFF9C4"})
        fmt_bar_wait = wb.add_format({"bg_color": "#FFE082", "border": 0})
        fmt_car    = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "bg_color": "#F5F5F5"})
        fmt_vline_blank = wb.add_format({"right": 2})  # 空白格右边框

        # 计算需要打竖线的列（右边框）
        vline_cols: set[int] = set()
        if primary_cycle and primary_cycle > 0:
            t = primary_cycle
            while t <= max_finish + 1e-9:
                col = 4 + int(math.ceil(t / grid_step)) - 1  # 右边界所在列
                if 4 <= col < 4 + n_cols_grid:
                    vline_cols.add(col)
                t += primary_cycle

        group_colors, _ = _palette()
        # 彩条格式缓存：key=(color_hex, with_right_border)
        bar_fmt_cache: Dict[tuple, Any] = {}
        wait_fmt_cache: Dict[bool, Any] = {False: fmt_bar_wait}

        def bar_fmt(group: str, display: str, with_right: bool):
            custom_hex = step_color_map.get(display)
            if not custom_hex:
                idx = (hash(group) >> 1) % len(group_colors)
                custom_hex = group_colors[idx]
            key = (custom_hex, with_right)
            fmt = bar_fmt_cache.get(key)
            if fmt is None:
                base = {"bg_color": custom_hex, "border": 0}
                if with_right:
                    base["right"] = 2
                fmt = wb.add_format(base)
                bar_fmt_cache[key] = fmt
            return fmt

        def wait_fmt(with_right: bool):
            fmt = wait_fmt_cache.get(with_right)
            if fmt is None:
                base = {"bg_color": "#FFE082", "border": 0, "right": 2}
                fmt = wb.add_format(base)
                wait_fmt_cache[with_right] = fmt
            return fmt

        # 列宽
        ws.set_column(0, 0, 36)
        ws.set_column(1, 1, 8)
        ws.set_column(2, 2, 18)
        ws.set_column(3, 3, 10)
        ws.set_column(4, 4 + n_cols_grid - 1, 2.8)

        # 表头
        ws.write(0, 0, f"连续投入{project}等待时间", fmt_header)
        ws.write(0, 1, "车号", fmt_header)
        ws.write(0, 2, "项目", fmt_header)
        ws.write(0, 3, "时间", fmt_header)
        for i in range(n_cols_grid):
            ws.write(0, 4 + i, f"{grid_step:.1f}", fmt_header)
        ws.freeze_panes(1, 0)

        first_data_row = 1
        row_cursor = 1

        for car in sorted(by_car.keys()):
            steps = by_car[car]
            if not steps:
                continue

            ewait = float(entry_wait.get(car, 0.0))
            twait = float(total_wait.get(car, 0.0))
            ws.write(row_cursor, 0, f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒", fmt_wait if ewait > 0 else fmt_left)
            ws.write(row_cursor, 1, car, fmt_car)
            ws.write(row_cursor, 2, "", fmt_left)
            ws.write(row_cursor, 3, _fmt_num(ewait) if ewait > 0 else "", fmt_text if ewait > 0 else fmt_left)

            if ewait > 0 and wait_policy == "before":
                first_start = steps[0]["start"]
                c0 = 4
                c1 = 4 + int(math.ceil(first_start / grid_step)) - 1
                c1 = max(c1, c0 - 1)
                for c in range(c0, c1 + 1):
                    use_right = c in vline_cols
                    ws.write(row_cursor, c, "", wait_fmt(use_right))
            row_cursor += 1

            for idx, s in enumerate(steps):
                # 左侧三列
                ws.write(row_cursor, 0, "", fmt_left)
                ws.write(row_cursor, 1, "", fmt_text)
                ws.write(row_cursor, 2, s["step_display"], fmt_left)
                ws.write(row_cursor, 3, _fmt_num(s["dur"]), fmt_text)

                # 服务条
                c_start = 4 + int(math.floor(s["start"] / grid_step))
                c_end_svc = 4 + int(math.ceil(s["svc_finish"] / grid_step)) - 1
                c_end_svc = max(c_end_svc, c_start)
                for c in range(c_start, c_end_svc + 1):
                    use_right = c in vline_cols
                    ws.write(row_cursor, c, "", bar_fmt(s["group"], s["step_display"], use_right))
                row_cursor += 1

                # 等待条（svc_finish → depart）
                if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                    wait_val = s["block_wait"]
                    next_name = steps[idx + 1]["step_display"]
                    ws.write(row_cursor, 0, f"等待{_fmt_num(wait_val)}秒（{s['step_display']} → {next_name}）", fmt_wait)
                    ws.write(row_cursor, 1, "", fmt_text)
                    ws.write(row_cursor, 2, "", fmt_wait)
                    ws.write(row_cursor, 3, _fmt_num(wait_val), fmt_text)
                    c_w0 = 4 + int(math.floor(s["svc_finish"] / grid_step))
                    c_w1 = 4 + int(math.ceil(s["depart"] / grid_step)) - 1
                    c_w1 = max(c_w1, c_w0)
                    for c in range(c_w0, c_w1 + 1):
                        use_right = c in vline_cols
                        ws.write(row_cursor, c, "", wait_fmt(use_right))
                    row_cursor += 1

            row_cursor += 1  # 车与车之间空一行

        last_data_row = row_cursor - 1

        # 空白格的贯穿竖线：让竖线“贯穿全表”，不受是否写入彩条影响
        if vline_cols and last_data_row >= first_data_row:
            for c in sorted(vline_cols):
                ws.conditional_format(first_data_row, c, last_data_row, c, {
                    "type": "blanks",
                    "format": fmt_vline_blank
                })


# ---------------- openpyxl 回退导出（文字） ---------------- #
def _export_with_openpyxl(rows: List[Dict[str, Any]], max_finish: float,
                           grid_step: float, wait_policy: str,
                           project: str, dst_path: str,
                           primary_cycle: float) -> None:
    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    out_rows = []
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        if not steps:
            continue
        ewait = float(entry_wait.get(car, 0.0))
        twait = float(total_wait.get(car, 0.0))
        out_rows.append({
            "车号": car,
            "项目": "(入站等待/总等待)",
            "时间": ewait,
            "说明": f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒"
        })
        for idx, s in enumerate(steps):
            out_rows.append({"车号": car, "项目": s["step_display"], "时间": s["dur"], "说明": ""})
            if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                out_rows.append({
                    "车号": car,
                    "项目": f"(等待：{s['step_display']}→{steps[idx+1]['step_display']})",
                    "时间": s["block_wait"],
                    "说明": f"等待{_fmt_num(s['block_wait'])}秒"
                })
        out_rows.append({"车号": "", "项目": "", "时间": "", "说明": ""})

    df = pd.DataFrame(out_rows, columns=["车号", "项目", "时间", "说明"])
    # 文字回退导出：不画彩条和网格线
    with pd.ExcelWriter(dst_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="作业组合票")
        try:
            ws = writer.sheets["作业组合票"]
            ws.freeze_panes = "A2"
        except Exception:
            pass