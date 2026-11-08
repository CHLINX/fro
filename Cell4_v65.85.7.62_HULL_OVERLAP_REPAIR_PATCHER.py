# -*- coding: utf-8 -*-
"""
FRO Cell4 v65.85.7.62 patcher
將 v65.85.7.59 原始來源轉成 v65.85.7.62：
1. 第一次盲腸排序只作暫存，不建立最終保護。
2. 凸包與原有跨日道路重複 guard 完成後，第二輪重新安排盲腸連續區塊。
3. 第二輪以正式 OSM scorer + 真實跨日物理道路重複量搜尋插入位置。
4. 通過工時、最近首站、Shared Snap、點位覆蓋、跨日重複與盲腸順序檢查後才鎖定。
"""

import os
import re
import json
import hashlib
import traceback
from pathlib import Path

OLD_VERSION = "v65.85.7.59"
NEW_VERSION = "v65.85.7.62"

_ip = get_ipython()
_history = _ip.user_ns.get("In", [])
_candidates = []
for _i, _cell in enumerate(_history):
    if not isinstance(_cell, str):
        continue
    if (
        len(_cell) > 1_000_000
        and f'CELL4_VERSION = "{OLD_VERSION}"' in _cell
        and f"[V{OLD_VERSION[1:]}_CONVEX_HULL_BALANCE]" in _cell
        and "def _v65_85_23_integrated_alns_split_optimizer" in _cell
    ):
        _candidates.append((_i, _cell))

if not _candidates:
    raise RuntimeError(
        "找不到完整 v65.85.7.59 Cell4 原始程式碼。請確認原始 Cell4 程式碼曾在目前 Colab 工作階段執行；"
        "本修補器不會猜測或改用其他版本。"
    )

_source_index, source = max(_candidates, key=lambda x: x[0])
if f'CELL4_VERSION = "{NEW_VERSION}"' in source:
    raise RuntimeError(f"來源已是 {NEW_VERSION}，拒絕重複套用。")

def _replace_once(text, old, new, label):
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f"{label}定位數量應為 1，實際為 {count}；停止修改。")
    return text.replace(old, new, 1)

# 版本號
source = _replace_once(
    source,
    f'CELL4_VERSION = "{OLD_VERSION}"',
    f'CELL4_VERSION = "{NEW_VERSION}"',
    "Cell4 版本號",
)


# ============================================================
# v65.85.7.62：凸包點位侵入診斷、全首站搜尋、群組聯合移動。
# 幾何目標採 Pareto 改善：點位侵入數與凸包交疊面積均不得惡化，
# 且至少一項必須改善；避免為降低一種問題而製造另一種問題。
# ============================================================

# 凸包移日重建時，不再把 12 個最近首站當作硬性搜尋上限。
source = _replace_once(
    source,
    "    def _v65_85_23_day_group_variants(groups):",
    "    def _v65_85_23_day_group_variants(groups, max_start_candidates=None):",
    "v62 首站候選參數",
)
source = _replace_once(
    source,
    """        _max_starts = min(len(_ranked_points), max(1, int(DAILY_NEAREST_START_MAX_CANDIDATES)))""",
    """        _start_cap = (
            int(DAILY_NEAREST_START_MAX_CANDIDATES)
            if max_start_candidates is None else int(max_start_candidates)
        )
        _max_starts = min(len(_ranked_points), max(1, _start_cap))""",
    "v62 首站候選上限可覆寫",
)
source = _replace_once(
    source,
    "    def _v65_85_23_evaluate_day_groups(groups, day_no, rescue_directions=True):",
    "    def _v65_85_23_evaluate_day_groups(groups, day_no, rescue_directions=True, max_start_candidates=None):",
    "v62 首站評估參數",
)
source = _replace_once(
    source,
    "        _variants, _nearest_point, _nearest_distance = _v65_85_23_day_group_variants(groups)",
    "        _variants, _nearest_point, _nearest_distance = _v65_85_23_day_group_variants(\n            groups, max_start_candidates=max_start_candidates\n        )",
    "v62 首站評估傳遞上限",
)
source = _replace_once(
    source,
    "        def _v65_85_30_local_group_2opt(batch, day_no):",
    "        def _v65_85_30_local_group_2opt(batch, day_no, exhaustive_start_candidates=False):",
    "v62 局部重排可要求完整首站搜尋",
)
source = _replace_once(
    source,
    """            _start_choice_30 = _v65_85_23_evaluate_day_groups(_raw_groups, int(day_no), rescue_directions=True)""",
    """            _start_choice_30 = _v65_85_23_evaluate_day_groups(
                _raw_groups, int(day_no), rescue_directions=True,
                max_start_candidates=(max(1, len(batch) - 1) if exhaustive_start_candidates else None),
            )""",
    "v62 局部重排傳入完整首站候選",
)
source = source.replace(
    "前12個最近首站均無法建立合規路線。",
    "所有已搜尋的最近首站候選均無法建立合規路線。",
)

# 凸包專用移日候選：在既有單組移動外，增加相鄰 2/3 群組整塊搬移與雙向一對一群組交換。
_candidates_anchor_62 = """            _candidate_rows.sort(key=lambda _r: ("""
_candidates_add_62 = r'''            # v62：既有單群組候選保留；額外建立相鄰 2/3 群組搬移與雙向交換。
            for _candidate_row_62 in _candidate_rows:
                _candidate_row_62.setdefault("operation", "MOVE_BLOCK")
                _candidate_row_62.setdefault("source_group_count", 1)

            _day_count_62 = len(_day_groups)
            for _src_idx_62 in range(_day_count_62):
                _src_groups_62 = _day_groups[_src_idx_62]
                _src_day_62 = _src_idx_62 + 1
                _src_point_count_62 = sum(len(_g) for _g in _src_groups_62)
                if len(_src_groups_62) < 2:
                    continue
                for _dst_idx_62 in range(_day_count_62):
                    if _dst_idx_62 == _src_idx_62:
                        continue
                    _dst_groups_62 = _day_groups[_dst_idx_62]
                    if not _dst_groups_62:
                        continue
                    _dst_day_62 = _dst_idx_62 + 1
                    _dst_point_count_62 = sum(len(_g) for _g in _dst_groups_62)
                    _src_service_before_62 = float(daily_service_minutes(
                        _src_day_62, _src_point_count_62
                    ))
                    _dst_service_before_62 = float(daily_service_minutes(
                        _dst_day_62, _dst_point_count_62
                    ))

                    # 相鄰群組作為不可拆散的聯合搬移單位。
                    for _block_count_62 in (2, 3):
                        if len(_src_groups_62) <= _block_count_62:
                            continue
                        for _src_pos_62 in range(len(_src_groups_62) - _block_count_62 + 1):
                            _block_groups_62 = [
                                list(_g) for _g in
                                _src_groups_62[_src_pos_62:_src_pos_62 + _block_count_62]
                            ]
                            _moving_points_62 = [
                                int(_p) for _g in _block_groups_62 for _p in _g
                            ]
                            _moving_count_62 = len(_moving_points_62)
                            if not _moving_count_62 or _moving_count_62 >= _src_point_count_62:
                                continue
                            _src_service_after_62 = float(daily_service_minutes(
                                _src_day_62, _src_point_count_62 - _moving_count_62
                            ))
                            _src_rough_62 = (
                                float(baseline_times[_src_idx_62])
                                + _src_service_after_62 - _src_service_before_62
                            )
                            for _insert_pos_62 in range(len(_dst_groups_62) + 1):
                                _dst_service_after_62 = float(daily_service_minutes(
                                    _dst_day_62, _dst_point_count_62 + _moving_count_62
                                ))
                                _rough_times_62 = [float(x) for x in baseline_times]
                                _rough_times_62[_src_idx_62] = _src_rough_62
                                _rough_times_62[_dst_idx_62] = (
                                    float(baseline_times[_dst_idx_62])
                                    + _dst_service_after_62 - _dst_service_before_62
                                )
                                _candidate_rows.append({
                                    "operation": "MOVE_BLOCK",
                                    "source_day": int(_src_day_62),
                                    "destination_day": int(_dst_day_62),
                                    "source_group_pos0": int(_src_pos_62),
                                    "source_group_count": int(_block_count_62),
                                    "destination_insert_group_pos0": int(_insert_pos_62),
                                    "moving_points": list(_moving_points_62),
                                    "estimated_daily_total_minutes": [
                                        round(float(x), 3) for x in _rough_times_62
                                    ],
                                    "estimated_max_daily_minutes": float(max(_rough_times_62)),
                                    "estimated_sum_daily_minutes": float(sum(_rough_times_62)),
                                })

            # 雙向一對一交換：每一對工作日只生成一次，避免方向重複。
            for _src_idx_62 in range(_day_count_62):
                _src_groups_62 = _day_groups[_src_idx_62]
                _src_day_62 = _src_idx_62 + 1
                _src_point_count_62 = sum(len(_g) for _g in _src_groups_62)
                if not _src_groups_62:
                    continue
                for _dst_idx_62 in range(_src_idx_62 + 1, _day_count_62):
                    _dst_groups_62 = _day_groups[_dst_idx_62]
                    if not _dst_groups_62:
                        continue
                    _dst_day_62 = _dst_idx_62 + 1
                    _dst_point_count_62 = sum(len(_g) for _g in _dst_groups_62)
                    for _src_pos_62, _src_group_62 in enumerate(_src_groups_62):
                        for _dst_pos_62, _dst_group_62 in enumerate(_dst_groups_62):
                            if set(_src_group_62) == set(_dst_group_62):
                                continue
                            _src_new_count_62 = (
                                _src_point_count_62 - len(_src_group_62) + len(_dst_group_62)
                            )
                            _dst_new_count_62 = (
                                _dst_point_count_62 - len(_dst_group_62) + len(_src_group_62)
                            )
                            _rough_times_62 = [float(x) for x in baseline_times]
                            _rough_times_62[_src_idx_62] += (
                                float(daily_service_minutes(_src_day_62, _src_new_count_62))
                                - float(daily_service_minutes(_src_day_62, _src_point_count_62))
                            )
                            _rough_times_62[_dst_idx_62] += (
                                float(daily_service_minutes(_dst_day_62, _dst_new_count_62))
                                - float(daily_service_minutes(_dst_day_62, _dst_point_count_62))
                            )
                            _candidate_rows.append({
                                "operation": "SWAP_GROUPS",
                                "source_day": int(_src_day_62),
                                "destination_day": int(_dst_day_62),
                                "source_group_pos0": int(_src_pos_62),
                                "source_group_count": 1,
                                "destination_group_pos0": int(_dst_pos_62),
                                "destination_insert_group_pos0": int(_dst_pos_62),
                                "moving_points": [int(_p) for _p in _src_group_62],
                                "swap_points": [int(_p) for _p in _dst_group_62],
                                "estimated_daily_total_minutes": [
                                    round(float(x), 3) for x in _rough_times_62
                                ],
                                "estimated_max_daily_minutes": float(max(_rough_times_62)),
                                "estimated_sum_daily_minutes": float(sum(_rough_times_62)),
                            })

'''
source = _replace_once(
    source, _candidates_anchor_62,
    _candidates_add_62 + _candidates_anchor_62,
    "v62 聯合群組搬移／交換候選",
)

# v62：凸包狀態同時統計對方點位侵入，以及「交集區內完全沒有任一日點位」的空白交疊。
source = _replace_once(
    source,
    """                from pyproj import Transformer
                from shapely.geometry import MultiPoint""",
    """                from pyproj import Transformer
                from shapely.geometry import MultiPoint, Point""",
    "v62 Shapely Point 匯入",
)
source = _replace_once(
    source,
    """                    _hulls = []
                    _point_day_map = {}
                    for _day_i, _batch in enumerate(_day_batches, start=1):
                        _xy = []""",
    """                    _hulls = []
                    _point_day_map = {}
                    _day_point_geometries = []
                    for _day_i, _batch in enumerate(_day_batches, start=1):
                        _xy = []
                        _day_points = []""",
    "v62 每日幾何點快取",
)
source = _replace_once(
    source,
    """                            _xy.append((_x_i, _y_i))
                            _point_day_map[_point_i] = int(_day_i)
                        _hull = MultiPoint(_xy).convex_hull if _xy else None
                        _hulls.append(_hull)""",
    """                            _xy.append((_x_i, _y_i))
                            _day_points.append((_point_i, Point(_x_i, _y_i)))
                            _point_day_map[_point_i] = int(_day_i)
                        _hull = MultiPoint(_xy).convex_hull if _xy else None
                        _hulls.append(_hull)
                        _day_point_geometries.append(_day_points)""",
    "v62 每日點位幾何保存",
)

_old_hull_pair_block_62 = """                    _overlap_rows = []
                    _total_overlap = 0.0
                    for _i in range(len(_hulls)):
                        _a = _hulls[_i]
                        if _a is None or _a.geom_type != "Polygon":
                            continue
                        for _j in range(_i + 1, len(_hulls)):
                            _b = _hulls[_j]
                            if _b is None or _b.geom_type != "Polygon":
                                continue
                            _area = float(_a.intersection(_b).area)
                            if _area > float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2):
                                _total_overlap += _area
                                _overlap_rows.append({
                                    "day_a": int(_i + 1), "day_b": int(_j + 1),
                                    "intersection_area_m2": round(_area, 3),
                                })
                    return {
                        "hulls": _hulls, "overlap_area_m2": float(_total_overlap),
                        "overlap_pairs": _overlap_rows, "point_day": _point_day_map,
                    }"""
_new_hull_pair_block_62 = r'''                    _overlap_rows = []
                    _intrusion_rows = []
                    _empty_overlap_rows = []
                    _total_overlap = 0.0
                    _total_intrusions = 0
                    _empty_overlap_area = 0.0
                    for _i in range(len(_hulls)):
                        _a = _hulls[_i]
                        if _a is None or _a.geom_type != "Polygon":
                            continue
                        for _j in range(_i + 1, len(_hulls)):
                            _b = _hulls[_j]
                            if _b is None or _b.geom_type != "Polygon":
                                continue
                            _intersection = _a.intersection(_b)
                            _area = float(_intersection.area)

                            _a_points_in_b = []
                            for _point_idx, _point_geom in _day_point_geometries[_i]:
                                if _b.covers(_point_geom):
                                    _relation = "INTERIOR" if _b.contains(_point_geom) else "BOUNDARY"
                                    _a_points_in_b.append({
                                        "point_index": int(_point_idx),
                                        "point": _v65_85_20_display_name(int(_point_idx)),
                                        "relation": _relation,
                                    })
                            _b_points_in_a = []
                            for _point_idx, _point_geom in _day_point_geometries[_j]:
                                if _a.covers(_point_geom):
                                    _relation = "INTERIOR" if _a.contains(_point_geom) else "BOUNDARY"
                                    _b_points_in_a.append({
                                        "point_index": int(_point_idx),
                                        "point": _v65_85_20_display_name(int(_point_idx)),
                                        "relation": _relation,
                                    })

                            _intersection_point_count = sum(
                                1 for _point_idx, _point_geom in (
                                    _day_point_geometries[_i] + _day_point_geometries[_j]
                                )
                                if _intersection.covers(_point_geom)
                            )
                            _pair_intrusion_count = len(_a_points_in_b) + len(_b_points_in_a)
                            _total_intrusions += int(_pair_intrusion_count)

                            _pair_row = {
                                "day_a": int(_i + 1),
                                "day_b": int(_j + 1),
                                "intersection_area_m2": round(_area, 3),
                                "day_a_points_inside_day_b_hull": int(len(_a_points_in_b)),
                                "day_b_points_inside_day_a_hull": int(len(_b_points_in_a)),
                                "point_intrusion_count": int(_pair_intrusion_count),
                                "intersection_contains_points": int(_intersection_point_count),
                                "empty_overlap_region": bool(
                                    _area > float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2)
                                    and _intersection_point_count == 0
                                ),
                                "day_a_intrusion_points": _a_points_in_b,
                                "day_b_intrusion_points": _b_points_in_a,
                            }
                            if (
                                _area > float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2)
                                or _pair_intrusion_count > 0
                            ):
                                _overlap_rows.append(_pair_row)
                            if _area > float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2):
                                _total_overlap += _area
                            if _pair_intrusion_count:
                                _intrusion_rows.append({
                                    "day_a": int(_i + 1),
                                    "day_b": int(_j + 1),
                                    "day_a_points_inside_day_b_hull": _a_points_in_b,
                                    "day_b_points_inside_day_a_hull": _b_points_in_a,
                                })
                            if _pair_row["empty_overlap_region"]:
                                _empty_overlap_area += _area
                                _empty_overlap_rows.append({
                                    "day_a": int(_i + 1), "day_b": int(_j + 1),
                                    "intersection_area_m2": round(_area, 3),
                                    "intersection_contains_points": 0,
                                })
                    return {
                        "hulls": _hulls,
                        "overlap_area_m2": float(_total_overlap),
                        "overlap_pairs": _overlap_rows,
                        "point_day": _point_day_map,
                        "point_intrusion_count": int(_total_intrusions),
                        "point_intrusion_rows": _intrusion_rows,
                        "empty_overlap_area_m2": float(_empty_overlap_area),
                        "empty_overlap_pairs": _empty_overlap_rows,
                    }'''
source = _replace_once(
    source, _old_hull_pair_block_62, _new_hull_pair_block_62,
    "v62 凸包交集與點位侵入完整診斷",
)

# 初始診斷寫入 log。
source = _replace_once(
    source,
    """                _convex_hull_balance_log_59["before_overlap_pairs"] = list(
                    _hull_state_59["overlap_pairs"]
                )
                _convex_hull_balance_log_59["status"] = (""",
    """                _convex_hull_balance_log_59["before_overlap_pairs"] = list(
                    _hull_state_59["overlap_pairs"]
                )
                _convex_hull_balance_log_59["before_point_intrusion_count"] = int(
                    _hull_state_59.get("point_intrusion_count", 0)
                )
                _convex_hull_balance_log_59["before_point_intrusion_rows"] = list(
                    _hull_state_59.get("point_intrusion_rows", [])
                )
                _convex_hull_balance_log_59["before_empty_overlap_area_m2"] = round(
                    float(_hull_state_59.get("empty_overlap_area_m2", 0.0)), 3
                )
                _convex_hull_balance_log_59["before_empty_overlap_pairs"] = list(
                    _hull_state_59.get("empty_overlap_pairs", [])
                )
                _convex_hull_balance_log_59["status"] = (""",
    "v62 初始幾何診斷 log",
)

source = _replace_once(
    source,
    """                _convex_hull_balance_log_59["status"] = (
                    "NO_POSITIVE_AREA_OVERLAP" if _hull_state_59["overlap_area_m2"]
                    <= float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2) else "SEARCHING"
                )""",
    """                _convex_hull_balance_log_59["status"] = (
                    "NO_POSITIVE_AREA_OVERLAP_AND_NO_POINT_INTRUSION"
                    if (
                        _hull_state_59["overlap_area_m2"]
                        <= float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2)
                        and int(_hull_state_59.get("point_intrusion_count", 0)) == 0
                    ) else "SEARCHING"
                )""",
    "v62 初始狀態同時檢查點位侵入",
)

# 幾何候選：只接受「點位侵入數不增加、凸包交疊面積不增加」且至少一項改善。
source = _replace_once(
    source,
    """                        _area_gain_59 = _base_area_59 - float(_trial_hull_59["overlap_area_m2"])
                        if _area_gain_59 > max(0.01, float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2) * 0.01):
                            _geometric_candidates_59.append((
                                -float(_area_gain_59),
                                float(_candidate_59.get("estimated_max_daily_minutes", float("inf"))),
                                float(_candidate_59.get("estimated_sum_daily_minutes", float("inf"))),
                                _candidate_59, _trial_hull_59,
                            ))
                    _geometric_candidates_59.sort(key=lambda _x: (_x[0], _x[1], _x[2]))""",
    """                        _area_gain_59 = _base_area_59 - float(_trial_hull_59["overlap_area_m2"])
                        _base_intrusions_62 = int(_hull_state_59.get("point_intrusion_count", 0))
                        _trial_intrusions_62 = int(_trial_hull_59.get("point_intrusion_count", 0))
                        _area_not_worse_62 = (
                            float(_trial_hull_59["overlap_area_m2"]) <= _base_area_59 + 1e-6
                        )
                        _intrusions_not_worse_62 = (
                            _trial_intrusions_62 <= _base_intrusions_62
                        )
                        _area_improved_62 = _area_gain_59 > max(
                            0.01, float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2) * 0.01
                        )
                        _intrusions_improved_62 = (
                            _trial_intrusions_62 < _base_intrusions_62
                        )
                        if (
                            _area_not_worse_62 and _intrusions_not_worse_62
                            and (_area_improved_62 or _intrusions_improved_62)
                        ):
                            _geometric_candidates_59.append((
                                -float(_area_gain_59),
                                float(_candidate_59.get("estimated_max_daily_minutes", float("inf"))),
                                float(_candidate_59.get("estimated_sum_daily_minutes", float("inf"))),
                                _candidate_59, _trial_hull_59,
                            ))
                    _geometric_candidates_59.sort(key=lambda _x: (
                        int(_x[4].get("point_intrusion_count", 0)),
                        float(_x[4]["overlap_area_m2"]),
                        float(_x[1]), float(_x[2]),
                    ))""",
    "v62 幾何候選 Pareto 篩選與排序",
)

# Search completeness metrics and route-attempt accounting.
source = _replace_once(
    source,
    """                        "exact_candidates_checked": 0,
                        "hard_constraint_rejections": 0,
                        "accepted": False,""",
    """                        "exact_candidates_checked": 0,
                        "route_reoptimization_attempts": 0,
                        "route_reoptimization_failures": 0,
                        "hard_constraint_rejections": 0,
                        "baseline_point_intrusion_count": int(
                            _hull_state_59.get("point_intrusion_count", 0)
                        ),
                        "geometric_point_intrusion_count_best": (
                            int(_geometric_candidates_59[0][4].get("point_intrusion_count", 0))
                            if _geometric_candidates_59 else None
                        ),
                        "accepted": False,""",
    "v62 幾何搜尋與路線嘗試統計",
)

# 幾何候選套用：支援 MOVE_BLOCK（1/2/3 群組）與 SWAP_GROUPS（雙向交換）。
_old_geom_apply_62 = """                        _moving_group_59 = _src_groups_59.pop(_src_pos_59)
                        if sorted(_moving_group_59) != sorted(int(p) for p in _candidate_59["moving_points"]):
                            continue
                        _dst_groups_59.insert(_dst_pos_59, list(_moving_group_59))"""
_new_geom_apply_62 = r'''                        _operation_62 = str(_candidate_59.get("operation", "MOVE_BLOCK"))
                        if _operation_62 == "SWAP_GROUPS":
                            _swap_pos_62 = int(_candidate_59["destination_group_pos0"])
                            if (
                                _src_pos_59 >= len(_src_groups_59)
                                or _swap_pos_62 < 0 or _swap_pos_62 >= len(_dst_groups_59)
                            ):
                                continue
                            _moving_group_59 = list(_src_groups_59[_src_pos_59])
                            _swap_group_62 = list(_dst_groups_59[_swap_pos_62])
                            if (
                                sorted(_moving_group_59) != sorted(int(p) for p in _candidate_59["moving_points"])
                                or sorted(_swap_group_62) != sorted(int(p) for p in _candidate_59["swap_points"])
                            ):
                                continue
                            _src_groups_59[_src_pos_59] = list(_swap_group_62)
                            _dst_groups_59[_swap_pos_62] = list(_moving_group_59)
                        else:
                            _source_count_62 = max(
                                1, int(_candidate_59.get("source_group_count", 1))
                            )
                            if _src_pos_59 + _source_count_62 > len(_src_groups_59):
                                continue
                            _moving_groups_62 = [
                                list(_g) for _g in
                                _src_groups_59[_src_pos_59:_src_pos_59 + _source_count_62]
                            ]
                            _moving_group_59 = [
                                int(_p) for _g in _moving_groups_62 for _p in _g
                            ]
                            if sorted(_moving_group_59) != sorted(int(p) for p in _candidate_59["moving_points"]):
                                continue
                            del _src_groups_59[_src_pos_59:_src_pos_59 + _source_count_62]
                            _dst_groups_59[_dst_pos_59:_dst_pos_59] = _moving_groups_62'''
source = _replace_once(
    source, _old_geom_apply_62, _new_geom_apply_62,
    "v62 幾何候選套用支援搬移與交換",
)

_old_exact_apply_62 = """                        _moving_group_59 = _src_groups_59.pop(_src_pos_59)
                        _dst_groups_59.insert(_dst_pos_59, list(_moving_group_59))"""
_new_exact_apply_62 = r'''                        _operation_62 = str(_candidate_59.get("operation", "MOVE_BLOCK"))
                        if _operation_62 == "SWAP_GROUPS":
                            _swap_pos_62 = int(_candidate_59["destination_group_pos0"])
                            if (
                                _src_pos_59 >= len(_src_groups_59)
                                or _swap_pos_62 < 0 or _swap_pos_62 >= len(_dst_groups_59)
                            ):
                                continue
                            _moving_group_59 = list(_src_groups_59[_src_pos_59])
                            _swap_group_62 = list(_dst_groups_59[_swap_pos_62])
                            if (
                                sorted(_moving_group_59) != sorted(int(p) for p in _candidate_59["moving_points"])
                                or sorted(_swap_group_62) != sorted(int(p) for p in _candidate_59["swap_points"])
                            ):
                                continue
                            _src_groups_59[_src_pos_59] = list(_swap_group_62)
                            _dst_groups_59[_swap_pos_62] = list(_moving_group_59)
                        else:
                            _source_count_62 = max(
                                1, int(_candidate_59.get("source_group_count", 1))
                            )
                            if _src_pos_59 + _source_count_62 > len(_src_groups_59):
                                continue
                            _moving_groups_62 = [
                                list(_g) for _g in
                                _src_groups_59[_src_pos_59:_src_pos_59 + _source_count_62]
                            ]
                            _moving_group_59 = [
                                int(_p) for _g in _moving_groups_62 for _p in _g
                            ]
                            if sorted(_moving_group_59) != sorted(int(p) for p in _candidate_59["moving_points"]):
                                continue
                            del _src_groups_59[_src_pos_59:_src_pos_59 + _source_count_62]
                            _dst_groups_59[_dst_pos_59:_dst_pos_59] = _moving_groups_62'''
source = _replace_once(
    source, _old_exact_apply_62, _new_exact_apply_62,
    "v62 精確路線候選套用支援搬移與交換",
)

# 凸包精確重建改用全部候選首站；不再因前 12 名均不可行就直接丟棄。
source = _replace_once(
    source,
    """                                    _v65_85_30_local_group_2opt(
                                        _trial_raw_batches_59[_affected_idx_59],
                                        _affected_idx_59 + 1,
                                    )""",
    """                                    _v65_85_30_local_group_2opt(
                                        _trial_raw_batches_59[_affected_idx_59],
                                        _affected_idx_59 + 1,
                                        exhaustive_start_candidates=True,
                                    )""",
    "v62 凸包候選完整首站重建",
)
source = _replace_once(
    source,
    """                            _pass_log_59["exact_candidates_checked"] += 1
                            if _trial_plan_score_59 is None:""",
    """                            _pass_log_59["exact_candidates_checked"] += 1
                            if _trial_plan_score_59 is None:""",
    "v62 精確評分計數定位",
)
# 每個候選進入正式路線重建前即計數；例外時不再出現「候選檢查 0」的誤導。
source = _replace_once(
    source,
    """                        _trial_batches_59 = [list(map(int, b)) for b in _base_batches_59]
                        try:""",
    """                        _trial_batches_59 = [list(map(int, b)) for b in _base_batches_59]
                        _pass_log_59["route_reoptimization_attempts"] += 1
                        try:""",
    "v62 正式路線重建嘗試計數",
)
source = _replace_once(
    source,
    """                        except Exception as _candidate_error_59:
                            _pass_log_59.setdefault("candidate_errors", []).append(""",
    """                        except Exception as _candidate_error_59:
                            _pass_log_59["route_reoptimization_failures"] += 1
                            _pass_log_59.setdefault("candidate_errors", []).append(""",
    "v62 正式路線重建失敗計數",
)

# 正式 scorer 回算後仍要求兩種幾何問題都不惡化，至少改善一項。
source = _replace_once(
    source,
    """                            if not (_exact_area_59 < _base_area_59 - max(
                                    0.01, float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2) * 0.01)):
                                continue""",
    """                            _base_intrusions_exact_62 = int(
                                _hull_state_59.get("point_intrusion_count", 0)
                            )
                            _exact_intrusions_62 = int(
                                _exact_hull_59.get("point_intrusion_count", 0)
                            )
                            _area_tolerance_exact_62 = max(
                                0.01, float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2) * 0.01
                            )
                            _area_nonworse_exact_62 = (
                                _exact_area_59 <= _base_area_59 + 1e-6
                            )
                            _intrusions_nonworse_exact_62 = (
                                _exact_intrusions_62 <= _base_intrusions_exact_62
                            )
                            _area_better_exact_62 = (
                                _exact_area_59 < _base_area_59 - _area_tolerance_exact_62
                            )
                            _intrusions_better_exact_62 = (
                                _exact_intrusions_62 < _base_intrusions_exact_62
                            )
                            if not (
                                _area_nonworse_exact_62 and _intrusions_nonworse_exact_62
                                and (_area_better_exact_62 or _intrusions_better_exact_62)
                            ):
                                continue""",
    "v62 精確幾何 Pareto 驗收",
)
source = _replace_once(
    source,
    """                            _key_59 = (
                                round(_exact_area_59, 3), tuple(_trial_plan_score_59),""",
    """                            _key_59 = (
                                int(_exact_intrusions_62),
                                round(_exact_area_59, 3), tuple(_trial_plan_score_59),""",
    "v62 精確候選以侵入數及面積排序",
)
source = _replace_once(
    source,
    """                        "moved_points": [
                            _v65_85_20_display_name(int(p))
                            for p in _accepted_move_59["moving_points"]
                        ],
                        "overlap_area_before_m2": round(_before_hull_59, 3),""",
    """                        "operation": str(_accepted_move_59.get("operation", "MOVE_BLOCK")),
                        "moved_points": [
                            _v65_85_20_display_name(int(p))
                            for p in _accepted_move_59["moving_points"]
                        ],
                        "swapped_points": [
                            _v65_85_20_display_name(int(p))
                            for p in _accepted_move_59.get("swap_points", [])
                        ],
                        "overlap_area_before_m2": round(_before_hull_59, 3),
                        "point_intrusions_before": int(
                            _convex_hull_balance_log_59["passes"][-1].get(
                                "baseline_point_intrusion_count", 0
                            )
                        ),
                        "point_intrusions_after": int(
                            _hull_state_59.get("point_intrusion_count", 0)
                        ),""",
    "v62 接受移動輸出詳細診斷",
)

# 最終凸包狀態加入點位侵入和空白交疊統計。
source = _replace_once(
    source,
    """                _convex_hull_balance_log_59["after_overlap_pairs"] = list(
                    _hull_state_59["overlap_pairs"]
                )
                if _hull_state_59["overlap_area_m2"] <= float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2):""",
    """                _convex_hull_balance_log_59["after_overlap_pairs"] = list(
                    _hull_state_59["overlap_pairs"]
                )
                _convex_hull_balance_log_59["after_point_intrusion_count"] = int(
                    _hull_state_59.get("point_intrusion_count", 0)
                )
                _convex_hull_balance_log_59["after_point_intrusion_rows"] = list(
                    _hull_state_59.get("point_intrusion_rows", [])
                )
                _convex_hull_balance_log_59["after_empty_overlap_area_m2"] = round(
                    float(_hull_state_59.get("empty_overlap_area_m2", 0.0)), 3
                )
                _convex_hull_balance_log_59["after_empty_overlap_pairs"] = list(
                    _hull_state_59.get("empty_overlap_pairs", [])
                )
                if (
                    _hull_state_59["overlap_area_m2"] <= float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2)
                    and int(_hull_state_59.get("point_intrusion_count", 0)) == 0
                ): """,
    "v62 最終幾何診斷輸出",
)

source = source.replace(
    "[V65.85.7.59_CONVEX_HULL_PROGRESS]",
    "[V65.85.7.62_CONVEX_HULL_PROGRESS]",
)

# 將所有進度標記升至 v62，並在結尾追加明確的診斷摘要標記。
source = _replace_once(
    source,
    """        print("[V65.85.7.59_CONVEX_HULL_BALANCE] " + json.dumps(
            _convex_hull_balance_log_59, ensure_ascii=False, default=str
        ))""",
    """        print("[V65.85.7.62_CONVEX_HULL_BALANCE] " + json.dumps(
            _convex_hull_balance_log_59, ensure_ascii=False, default=str
        ), flush=True)
        print("[V65.85.7.62_HULL_GEOMETRY_DIAGNOSTIC] " + json.dumps({
            "before_point_intrusion_count": _convex_hull_balance_log_59.get("before_point_intrusion_count"),
            "after_point_intrusion_count": _convex_hull_balance_log_59.get("after_point_intrusion_count"),
            "before_overlap_area_m2": _convex_hull_balance_log_59.get("before_overlap_area_m2"),
            "after_overlap_area_m2": _convex_hull_balance_log_59.get("after_overlap_area_m2"),
            "before_empty_overlap_area_m2": _convex_hull_balance_log_59.get("before_empty_overlap_area_m2"),
            "after_empty_overlap_area_m2": _convex_hull_balance_log_59.get("after_empty_overlap_area_m2"),
            "before_point_intrusion_rows": _convex_hull_balance_log_59.get("before_point_intrusion_rows", []),
            "after_point_intrusion_rows": _convex_hull_balance_log_59.get("after_point_intrusion_rows", []),
            "after_empty_overlap_pairs": _convex_hull_balance_log_59.get("after_empty_overlap_pairs", []),
        }, ensure_ascii=False, default=str), flush=True)""",
    "v62 幾何診斷輸出標記",
)

# 第一次成功盲腸排序：保留實際 point index，供凸包／全域回退後的第二輪重新排列。
_old_accepted = '''                    "full_bilateral_order_after": [
                        _v65_85_42_point_label(int(_p_52))
                        for _p_52 in _required_branch_sequence_52
                    ],'''
_new_accepted = _old_accepted + '''
                    "full_bilateral_order_after_indices": [
                        int(_p_52) for _p_52 in _required_branch_sequence_52
                    ],'''
source = _replace_once(
    source, _old_accepted, _new_accepted,
    "已接受盲腸走廊 point-index 保存",
)

_old_already = '''                        "left_return_order": [
                            _v65_85_42_point_label(int(_p_52))
                            for _u_52 in _left_units_52 for _p_52 in _u_52["points"]
                        ],'''
_new_already = _old_already + '''
                        "full_bilateral_order_after_indices": [
                            int(_p_52) for _p_52 in _required_branch_sequence_52
                        ],'''
source = _replace_once(
    source, _old_already, _new_already,
    "原本正確盲腸走廊 point-index 保存",
)

# 在日別確定、全域分日／群組 guard 結束後鎖住「日別成員」，但不鎖日內順序。
_membership_anchor = '''    if VEHICLE_MODE == "MOTOR":
        _v65_85_46_audit_plan("AFTER_DAILY_TIME_REFINEMENT_AND_GLOBAL_CLUSTER_GUARDS", batches)'''
_membership_insert = _membership_anchor + '''
        # v65.85.7.62：此時凸包與所有跨日分日調整已結束；只鎖日別成員，不鎖第一次盲腸排序。
        _V65_85_61_POST_HULL_DAY_MEMBERSHIP = tuple(
            tuple(sorted(int(p) for p in _b[1:])) for _b in batches
        )
        _V65_85_61_POST_HULL_MEMBERSHIP_SHA256 = hashlib.sha256(
            json.dumps(_V65_85_61_POST_HULL_DAY_MEMBERSHIP, separators=(",", ":")).encode("utf-8")
        ).hexdigest()'''
source = _replace_once(
    source, _membership_anchor, _membership_insert,
    "凸包完成後日別成員快照",
)

# 第二輪盲腸排序：放在第一次日內最佳化與跨日道路回退 guard 之後。
# 對每條走廊搜尋多個可插入位置；beam search 同時考慮最終跨日實際物理道路重複量。
_final_anchor = '''        _validate_cross_day_batches_v65_14(batches)
        _v65_85_20_final_coverage_check(
            batches, _v65_85_37_expected_batches_for_final_check
        )'''

_second_pass = r'''        # ============================================================
        # v65.85.7.62：第二輪盲腸排序（凸包／跨日 guard 完成後）
        # 第一次 CBA 只保存走廊定序，不保護日內位置；此處才做最終排序與鎖定。
        # 所有候選都使用正式 OSM scorer；以實際跨日物理道路重複量作第一排序目標。
        # ============================================================
        _v65_85_61_base_batches = [list(map(int, _b)) for _b in batches]
        _v65_85_61_base_overlap = _v65_85_55_cross_day_road_overlap(
            _v65_85_61_base_batches
        )
        _v65_85_61_overlap_limit_m = float(
            _v65_85_56_pre_local_overlap.get(
                "repeat_road_m",
                _v65_85_61_base_overlap.get("repeat_road_m", 0.0),
            )
        ) + 0.05

        _v65_85_61_constraints = []
        _v65_85_61_point_owner = {}
        for _d61, _stat61 in enumerate(_v65_85_42_day_stats, start=1):
            _sweep61 = _stat61.get("dead_end_right_side_sweep", {}) or {}
            _rows61 = list(_sweep61.get("accepted_corridors", []))
            _rows61 += [
                _r61 for _r61 in _sweep61.get("skipped_corridors", [])
                if str(_r61.get("reason", "")) == "BILATERAL_ORDER_ALREADY_CORRECT"
            ]
            _day_set61 = set(map(int, _v65_85_61_base_batches[_d61 - 1][1:]))
            for _row61 in _rows61:
                _seq61 = [int(p) for p in _row61.get("full_bilateral_order_after_indices", [])]
                if len(_seq61) < 2 or len(_seq61) != len(set(_seq61)):
                    raise RuntimeError(
                        f"v65.85.7.62：盲腸走廊 {_row61.get('branch_id')} 缺少有效完整 point-index 序列。"
                    )
                if not set(_seq61).issubset(_day_set61):
                    raise RuntimeError(
                        f"v65.85.7.62：凸包分日後盲腸走廊 {_row61.get('branch_id')} 跨日或遺失，拒絕鎖定。"
                    )
                for _p61 in _seq61:
                    _owner61 = _v65_85_61_point_owner.get(_p61)
                    if _owner61 is not None and _owner61 != (_d61, str(_row61.get("branch_id"))):
                        raise RuntimeError(
                            f"v65.85.7.62：同一勘查點同時屬於多條盲腸走廊：{_p61}。"
                        )
                    _v65_85_61_point_owner[_p61] = (_d61, str(_row61.get("branch_id")))
                _positions61 = [
                    _v65_85_61_base_batches[_d61 - 1].index(_p61)
                    for _p61 in _seq61
                ]
                _v65_85_61_constraints.append({
                    "day_no": int(_d61),
                    "branch_id": str(_row61.get("branch_id", "")),
                    "sequence": _seq61,
                    "base_first_position": int(min(_positions61)),
                })

        # 同一天先處理原路線中位置較後的走廊，避免後續插入切開已完成的走廊區塊。
        _v65_85_61_constraints.sort(
            key=lambda _c61: (_c61["day_no"], -_c61["base_first_position"], _c61["branch_id"])
        )

        def _v65_85_61_contains_sequence(_values61, _target61):
            _n61 = len(_target61)
            if _n61 == 0 or _n61 > len(_values61):
                return False
            return any(
                _values61[_i61:_i61 + _n61] == _target61
                for _i61 in range(len(_values61) - _n61 + 1)
            )

        def _v65_85_61_plan_metrics(_plan61):
            _times61 = [
                float(actual_closed_day_metrics(_b61, _d61)["total_min"])
                for _d61, _b61 in enumerate(_plan61, start=1)
            ]
            _dist61 = [
                float(actual_closed_day_metrics(_b61, _d61).get("distance_m", 0.0))
                for _d61, _b61 in enumerate(_plan61, start=1)
            ]
            return _times61, _dist61

        _v65_85_61_beam = [{
            "batches": [list(map(int, _b61)) for _b61 in _v65_85_61_base_batches],
            "protected": [[] for _ in _v65_85_61_base_batches],
            "overlap_m": float(_v65_85_61_base_overlap.get("repeat_road_m", 0.0)),
        }]
        _v65_85_61_beam_width = 4
        _v65_85_61_local_shortlist = 3
        _v65_85_61_candidate_evaluations = 0
        _v65_85_61_overlap_evaluations = 0
        _v65_85_61_rejections = {
            "route_infeasible": 0,
            "daily_cap_or_nearest_rank": 0,
            "shared_snap": 0,
            "endpoint_or_insertion": 0,
        }

        for _ci61, _constraint61 in enumerate(_v65_85_61_constraints, start=1):
            _new_beam61 = []
            _day_idx61 = int(_constraint61["day_no"]) - 1
            _target61 = list(map(int, _constraint61["sequence"]))
            _target_set61 = set(_target61)
            _required_rank61 = int(
                _v65_85_42_day_stats[_day_idx61].get("final_nearest_start_rank", 1)
            )
            for _state61 in _v65_85_61_beam:
                _day61 = list(map(int, _state61["batches"][_day_idx61]))
                _visits61 = _day61[1:]
                if not _target_set61.issubset(set(_visits61)):
                    _v65_85_61_rejections["endpoint_or_insertion"] += 1
                    continue

                # 已經符合連續完整序列時，直接保留，不做多餘重排。
                if _v65_85_61_contains_sequence(_visits61, _target61):
                    _same61 = {
                        "batches": [list(map(int, _b61)) for _b61 in _state61["batches"]],
                        "protected": [list(_x61) for _x61 in _state61["protected"]],
                        "overlap_m": float(_state61["overlap_m"]),
                    }
                    _same61["protected"][_day_idx61].append(_target61)
                    _new_beam61.append(_same61)
                    continue

                _remaining61 = [int(p61) for p61 in _visits61 if int(p61) not in _target_set61]
                if len(_remaining61) < 2 or _target_set61.intersection({_day61[1], _day61[-1]}):
                    _v65_85_61_rejections["endpoint_or_insertion"] += 1
                    continue

                # 不允許新區塊插入已完成盲腸區塊的內部。
                _invalid_gaps61 = set()
                _protected_sequences61 = list(_state61["protected"][_day_idx61])
                for _prev61 in _protected_sequences61:
                    _prev_pos61 = next((
                        _k61 for _k61 in range(len(_remaining61) - len(_prev61) + 1)
                        if _remaining61[_k61:_k61 + len(_prev61)] == _prev61
                    ), None)
                    if _prev_pos61 is None:
                        raise RuntimeError(
                            "v65.85.7.62：先前完成的盲腸區塊被後續候選拆開，拒絕繼續。"
                        )
                    _invalid_gaps61.update(
                        range(_prev_pos61 + 1, _prev_pos61 + len(_prev61))
                    )

                _local_candidates61 = []
                for _gap61 in range(1, len(_remaining61)):
                    if _gap61 in _invalid_gaps61:
                        continue
                    _candidate_order61 = (
                        [_day61[0]]
                        + _remaining61[:_gap61]
                        + _target61
                        + _remaining61[_gap61:]
                    )
                    if (
                        _candidate_order61[1] != _day61[1]
                        or _candidate_order61[-1] != _day61[-1]
                        or len(_candidate_order61) != len(_day61)
                        or len(set(_candidate_order61[1:])) != len(_visits61)
                    ):
                        _v65_85_61_rejections["endpoint_or_insertion"] += 1
                        continue
                    if not _v65_85_46_shared_snap_order_status(_candidate_order61)[0]:
                        _v65_85_61_rejections["shared_snap"] += 1
                        continue

                    _v65_85_61_candidate_evaluations += 1
                    _score61 = _v65_85_23_score_order(
                        _candidate_order61, int(_constraint61["day_no"]), _visits61
                    )
                    if _score61 is None or not bool(_score61.get("feasible", False)):
                        _v65_85_61_rejections["route_infeasible"] += 1
                        continue
                    _total61 = float(_score61.get("total_min", float("inf")))
                    _rank61 = int(_score61.get("nearest_start_rank", 10**9))
                    if (
                        not np.isfinite(_total61)
                        or _total61 > float(max_day_min) + 1e-9
                        or _rank61 != _required_rank61
                    ):
                        _v65_85_61_rejections["daily_cap_or_nearest_rank"] += 1
                        continue
                    _local_candidates61.append((
                        _total61,
                        float(_score61.get("distance_m", float("inf"))),
                        int(_gap61),
                        _candidate_order61,
                    ))

                _local_candidates61.sort(key=lambda _x61: (_x61[0], _x61[1], _x61[2]))
                for _local61 in _local_candidates61[:_v65_85_61_local_shortlist]:
                    _trial61 = [
                        list(map(int, _b61)) for _b61 in _state61["batches"]
                    ]
                    _trial61[_day_idx61] = list(map(int, _local61[3]))
                    _overlap61 = _v65_85_55_cross_day_road_overlap(_trial61)
                    _v65_85_61_overlap_evaluations += 1
                    _new_protected61 = [list(_x61) for _x61 in _state61["protected"]]
                    _new_protected61[_day_idx61].append(_target61)
                    _new_beam61.append({
                        "batches": _trial61,
                        "protected": _new_protected61,
                        "overlap_m": float(_overlap61.get("repeat_road_m", float("inf"))),
                    })

            if not _new_beam61:
                raise RuntimeError(
                    f"v65.85.7.62：第二輪盲腸排序無可行候選；day={_constraint61['day_no']} "
                    f"branch={_constraint61['branch_id']}。"
                )

            _dedup61 = {}
            for _state61 in _new_beam61:
                _sig61 = tuple(tuple(int(p61) for p61 in _b61) for _b61 in _state61["batches"])
                _old61 = _dedup61.get(_sig61)
                if _old61 is None or _state61["overlap_m"] < _old61["overlap_m"]:
                    _dedup61[_sig61] = _state61
            _ranked_states61 = list(_dedup61.values())
            for _state61 in _ranked_states61:
                _times61, _distances61 = _v65_85_61_plan_metrics(_state61["batches"])
                _state61["sort_key"] = (
                    round(float(_state61["overlap_m"]), 3),
                    round(float(sum(_times61)), 3),
                    round(float(max(_times61) if _times61 else 0.0), 3),
                    round(float(sum(_distances61)), 2),
                )
            _ranked_states61.sort(key=lambda _s61: _s61["sort_key"])
            _v65_85_61_beam = _ranked_states61[:_v65_85_61_beam_width]
            print("[V65.85.7.62_POST_HULL_CBA_PROGRESS] " + json.dumps({
                "corridor_no": int(_ci61),
                "corridor_total": int(len(_v65_85_61_constraints)),
                "day_no": int(_constraint61["day_no"]),
                "branch_id": str(_constraint61["branch_id"]),
                "beam_states": int(len(_v65_85_61_beam)),
                "candidate_route_checks": int(_v65_85_61_candidate_evaluations),
                "global_overlap_checks": int(_v65_85_61_overlap_evaluations),
                "best_repeat_road_km": round(float(_v65_85_61_beam[0]["overlap_m"]) / 1000.0, 6),
            }, ensure_ascii=False, default=str), flush=True)

        _v65_85_61_valid_beam = [
            _s61 for _s61 in _v65_85_61_beam
            if float(_s61["overlap_m"]) <= float(_v65_85_61_overlap_limit_m)
        ]
        if not _v65_85_61_valid_beam:
            _best_repeat61 = (
                min(float(_s61["overlap_m"]) for _s61 in _v65_85_61_beam)
                if _v65_85_61_beam else float("inf")
            )
            raise RuntimeError(
                "v65.85.7.62：第二輪盲腸順序雖可重建，但沒有候選能同時維持原跨日道路重複上限；"
                f"baseline={_v65_85_61_overlap_limit_m/1000.0:.6f}km，"
                f"best={_best_repeat61/1000.0:.6f}km。拒絕輸出不合規路線。"
            )

        _v65_85_61_best = min(
            _v65_85_61_valid_beam,
            key=lambda _s61: _s61.get("sort_key", (float("inf"),)),
        )
        batches = [list(map(int, _b61)) for _b61 in _v65_85_61_best["batches"]]

        # 最終實際序列驗證：每條已確認走廊都必須整段連續、方向完全相同。
        _v65_85_61_final_corridors = []
        _v65_85_61_dead_end_failures = []
        for _constraint61 in _v65_85_61_constraints:
            _day61 = batches[int(_constraint61["day_no"]) - 1]
            _visits61 = _day61[1:]
            _ok61 = _v65_85_61_contains_sequence(
                _visits61, list(map(int, _constraint61["sequence"]))
            )
            _row61 = {
                "day_no": int(_constraint61["day_no"]),
                "branch_id": str(_constraint61["branch_id"]),
                "sequence": [
                    _v65_85_42_point_label(int(p61))
                    for p61 in _constraint61["sequence"]
                ],
                "pass": bool(_ok61),
            }
            _v65_85_61_final_corridors.append(_row61)
            if not _ok61:
                _v65_85_61_dead_end_failures.append(_row61)
        if _v65_85_61_dead_end_failures:
            raise RuntimeError(
                "v65.85.7.62：第二輪後仍有盲腸走廊未符合完整連續順序，停止交接。"
            )

        _v65_85_61_final_overlap = _v65_85_55_cross_day_road_overlap(batches)
        if (
            float(_v65_85_61_final_overlap.get("repeat_road_m", float("inf")))
            > float(_v65_85_61_overlap_limit_m) + 1e-6
        ):
            raise RuntimeError(
                "v65.85.7.62：最終跨日物理道路重複超過基準，停止交接。"
            )

        _v65_85_61_final_membership = tuple(
            tuple(sorted(int(p61) for p61 in _b61[1:])) for _b61 in batches
        )
        if _v65_85_61_final_membership != _V65_85_61_POST_HULL_DAY_MEMBERSHIP:
            raise RuntimeError(
                "v65.85.7.62：第二輪盲腸調整意外改變凸包完成後的日別成員，停止交接。"
            )

        _v65_85_61_final_times, _v65_85_61_final_distances = _v65_85_61_plan_metrics(batches)
        if any(
            (not np.isfinite(_t61)) or _t61 > float(max_day_min) + 1e-9
            for _t61 in _v65_85_61_final_times
        ):
            raise RuntimeError("v65.85.7.62：第二輪盲腸調整後有工作日超過工時上限。")

        # 凸包最終稽核：固定起點不計入；僅正面積交集算重疊。
        from pyproj import Transformer as _Transformer61
        from shapely.geometry import MultiPoint as _MultiPoint61
        _transformer61 = _Transformer61.from_crs("EPSG:4326", "EPSG:3826", always_xy=True)
        _hulls61 = []
        for _day61 in batches:
            _xy61 = []
            for _p61 in _day61[1:]:
                if int(_p61) == int(start_idx):
                    continue
                _lon61, _lat61 = map(float, coords[int(_p61)][:2])
                _x61, _y61 = _transformer61.transform(_lon61, _lat61)
                _xy61.append((_x61, _y61))
            _hulls61.append(_MultiPoint61(_xy61).convex_hull if _xy61 else None)
        _hull_pairs61 = []
        _hull_area61 = 0.0
        for _i61 in range(len(_hulls61)):
            if _hulls61[_i61] is None or _hulls61[_i61].geom_type != "Polygon":
                continue
            for _j61 in range(_i61 + 1, len(_hulls61)):
                if _hulls61[_j61] is None or _hulls61[_j61].geom_type != "Polygon":
                    continue
                _area61 = float(_hulls61[_i61].intersection(_hulls61[_j61]).area)
                if _area61 > float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2):
                    _hull_area61 += _area61
                    _hull_pairs61.append({
                        "day_a": int(_i61 + 1), "day_b": int(_j61 + 1),
                        "intersection_area_m2": round(_area61, 3),
                    })

        _V65_85_61_FINAL_DAY_MEMBERSHIP_LOCK = _v65_85_61_final_membership
        _V65_85_61_FINAL_SEQUENCE_LOCK = tuple(tuple(map(int, _b61)) for _b61 in batches)
        _V65_85_61_FINAL_SEQUENCE_SHA256 = hashlib.sha256(
            json.dumps(_V65_85_61_FINAL_SEQUENCE_LOCK, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        print("[V65.85.7.62_FINAL_AUDIT] " + json.dumps({
            "version": "v65.85.7.62",
            "first_dead_end_pass_protected": False,
            "second_dead_end_pass": "PASS_AND_LOCKED",
            "verified_corridor_count": int(len(_v65_85_61_final_corridors)),
            "verified_corridors": _v65_85_61_final_corridors,
            "day_membership_locked": True,
            "day_membership_sha256": _V65_85_61_POST_HULL_MEMBERSHIP_SHA256,
            "final_sequence_sha256": _V65_85_61_FINAL_SEQUENCE_SHA256,
            "daily_total_minutes": [round(float(x61), 3) for x61 in _v65_85_61_final_times],
            "daily_cap_minutes": float(max_day_min),
            "coverage_expected_points": int(sum(len(_b61) - 1 for _b61 in batches)),
            "cross_day_repeat_road_km": round(
                float(_v65_85_61_final_overlap.get("repeat_road_m", 0.0)) / 1000.0, 6
            ),
            "cross_day_repeat_limit_km": round(_v65_85_61_overlap_limit_m / 1000.0, 6),
            "shared_physical_segment_count": int(
                _v65_85_61_final_overlap.get("shared_physical_segment_count", 0)
            ),
            "final_convex_hull_overlap_area_m2": round(float(_hull_area61), 3),
            "final_convex_hull_overlap_pairs": _hull_pairs61,
            "convex_hull_status": (
                "NO_POSITIVE_AREA_OVERLAP"
                if _hull_area61 <= float(CONVEX_HULL_OVERLAP_AREA_TOLERANCE_M2)
                else "UNRESOLVED_POSITIVE_AREA_OVERLAP"
            ),
            "candidate_route_checks": int(_v65_85_61_candidate_evaluations),
            "global_overlap_checks": int(_v65_85_61_overlap_evaluations),
            "rejections": _v65_85_61_rejections,
        }, ensure_ascii=False, default=str), flush=True)

'''

source = _replace_once(
    source, _final_anchor, _second_pass + _final_anchor,
    "凸包後第二輪盲腸排序與最終鎖定插入點",
)

# 最終 Python compile 預檢；不會觸碰 GitHub main。
_output = Path("/content/Cell4_v65.85.7.62.py")
_compiled = compile(source, str(_output), "exec")
_output.write_text(source, encoding="utf-8", newline="\n")
_sha = hashlib.sha256(source.encode("utf-8")).hexdigest()

print("=" * 78)
print("巡田最速適 FRO｜Cell4 v65.85.7.62")
print("=" * 78)
print(f"來源：IPython In[{_source_index}]，{len(source):,} 字元")
print(f"輸出：{_output}")
print(f"大小：{_output.stat().st_size:,} bytes")
print(f"SHA256：{_sha}")
print("Python compile：PASS")
print("修改項目：第一次盲腸不鎖定 → 凸包／跨日 guard → 第二輪盲腸位置搜尋 → 最終鎖定")
print("開始執行完整 Cell4 ...")
print("=" * 78)

_namespace = _ip.user_ns
_original_file = _namespace.get("__file__")
try:
    _namespace["__file__"] = str(_output)
    exec(_compiled, _namespace, _namespace)
finally:
    if _original_file is None:
        _namespace.pop("__file__", None)
    else:
        _namespace["__file__"] = _original_file
