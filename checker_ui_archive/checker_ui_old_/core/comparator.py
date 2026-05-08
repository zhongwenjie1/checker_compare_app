import pandas as pd
import re, unicodedata

def _normalize_key(s: str) -> str:
    """BC POS NAME / Parts Name 等作 key 时：大小写、全角半角统一"""
    return unicodedata.normalize("NFKC", str(s)).strip().lower()


def _pn_key_alnum(p):
    """品番：保留字母数字，转大写，取前 5 位；用于兼容 8646C 这类写法"""
    s = re.sub(r'[^0-9A-Za-z]', '', str(p)).upper()
    return s[:5]

def _pn_keys_multi(p):
    """标准文件的品番可能用'/'分隔多个，只要任一匹配即 OK"""
    if pd.isna(p):
        return []
    keys = []
    for seg in str(p).split('/'):
        seg = seg.strip()
        if not seg:
            continue
        keys.append(_pn_key_alnum(seg))
    return keys or [""]

def _assy8(a):
    """任意字符串抽取 8 位组立番号（不足补 0）"""
    s = re.sub(r'\D', '', str(a))
    if len(s) >= 8:
        return s[:8]
    return s.zfill(8)

def _find_col(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None

def _split_std_assy_list(v: str) -> list[str]:
    """标准侧：'xxx/yyy/zzz' → [八位, 八位, ...]"""
    if pd.isna(v):
        return []
    outs = []
    for seg in str(v).split('/'):
        seg = seg.strip()
        if not seg:
            continue
        outs.append(_assy8(seg))
    return outs

def compare(std_df: pd.DataFrame, sys_df: pd.DataFrame,
            std_key_col='BC POS NAME', sys_key_col='BC POS NAME',
            std_pn='品番',        sys_pn='品番',
            std_as='组立番号',    sys_as='组立番号',
            upload_col='是否上传'):
    """
    修正版：
      1) key 做 NFKC+lower 统一，避免大小写/全角半角导致的错判
      2) 品番比较统一成 “只取数字前 5 位”
      3) 系统端/标准端组立番号都统一 8 位（标准端支持‘/’多值）
      4) 标准品番 = ALL ⇒ 组里只要有 上传=1 就判 OK
      5) 上传=0 ⇒ 一律“未比对”（灰色）
    """
    std_df = std_df.copy()
    sys_df = sys_df.copy()

    # 列名自动识别
    std_key_col = _find_col(std_df, [std_key_col, 'BC POS Name', 'BC POS', '零件名称', 'Parts Name', '__KEY__'])
    sys_key_col = _find_col(sys_df, [sys_key_col, 'BC POS Name', 'BC POS', '零件名称', 'Parts Name', '__KEY__'])
    upload_col  = _find_col(sys_df, [upload_col, '上传', '是否上传'])
    std_pn      = _find_col(std_df, [std_pn, '标准品番', 'PN', '品番（标准）'])
    sys_pn      = _find_col(sys_df, [sys_pn, '系统品番', 'PN', '品番（系统）'])
    std_as      = _find_col(std_df, [std_as, '标准组立番号'])
    sys_as      = _find_col(sys_df, [sys_as, '系统组立番号', 'GP.CP./HIKI. ITEM', '组立番号'])

    if not all([std_key_col, sys_key_col, upload_col, std_pn, sys_pn, std_as, sys_as]):
        raise ValueError("缺少必要列，无法比对（请检查“品番/组立番号/是否上传/BC POS NAME”等列名）")

    # 统一 key
    std_df['__KEY__'] = std_df[std_key_col].astype(str).map(_normalize_key)
    sys_df['__KEY__'] = sys_df[sys_key_col].astype(str).map(_normalize_key)

    # 结果列初始化
    std_df[['比对结果', 'NG原因']] = ['', '']
    sys_df[['比对结果', 'NG原因']] = ['未配对', '']

    # 上传标记
    upload0_mask = sys_df[upload_col].astype(str).str.strip() == '0'

    # 预计算系统端品番/组立
    sys_df['__pn5'] = sys_df[sys_pn].map(_pn_key_alnum)
    sys_df['__assy8'] = sys_df[sys_as].map(_assy8)

    # 预计算标准端品番/组立
    std_df['__pn5_list'] = std_df[std_pn].map(_pn_keys_multi)
    std_df['__assy_list'] = std_df[std_as].map(_split_std_assy_list)

    for idx, row in std_df.iterrows():
        key = row['__KEY__']
        pn_keys = row['__pn5_list']
        candidates = row['__assy_list']

        grp_all = sys_df[sys_df['__KEY__'] == key]
        grp1 = grp_all[~upload0_mask.loc[grp_all.index]]  # 只看上传=1

        # 标准品番=ALL
        if str(row[std_pn]).strip().upper() == 'ALL':
            if len(grp1) > 0:
                std_df.loc[idx, ['比对结果', 'NG原因']] = ['OK', '']
                sys_df.loc[grp1.index, ['比对结果', 'NG原因']] = ['OK', '标准品番=ALL']
            else:
                std_df.loc[idx, ['比对结果', 'NG原因']] = ['NG', '无上传=1 的系统行']
            continue

        matched = False
        for jdx, r in grp1.iterrows():
            if (r['__pn5'] in pn_keys) and (not candidates or r['__assy8'] in candidates):
                matched = True
                sys_df.loc[jdx, ['比对结果', 'NG原因']] = ['OK', '']

        if matched:
            std_df.loc[idx, ['比对结果', 'NG原因']] = ['OK', '']
            # 同 key 其它仍“未配对”的上传=1 行标 NG（可选）
            remaining = grp1[sys_df.loc[grp1.index, '比对结果'] == '未配对']
            if len(remaining) > 0:
                sys_df.loc[remaining.index, ['比对结果', 'NG原因']] = ['NG', '品番或组立不一致']
        else:
            std_df.loc[idx, ['比对结果', 'NG原因']] = ['NG', '品番或组立不一致']
            remaining = grp1[sys_df.loc[grp1.index, '比对结果'] == '未配对']
            if len(remaining) > 0:
                sys_df.loc[remaining.index, ['比对结果', 'NG原因']] = ['NG', '品番或组立不一致']

    # 上传=0 固定灰色“未比对”
    sys_df.loc[upload0_mask, ['比对结果', 'NG原因']] = ['未比对', '不上传']
    return std_df, sys_df