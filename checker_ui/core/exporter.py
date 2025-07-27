import pandas as pd
from datetime import datetime

def export(std_df: pd.DataFrame, sys_df: pd.DataFrame, path: str):
    """
    将标准 & 系统两个结果表导出到一个 Sheet（左右分区）
    - 首行冻结
    - OK/NG/未比对 三色底
    - 自动列宽
    - 末尾统计
    """
    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        book = writer.book
        sheet = '对比结果'

        std_df.to_excel(writer, sheet_name=sheet, startrow=0, startcol=0, index=False)
        sys_df.to_excel(writer, sheet_name=sheet, startrow=0, startcol=len(std_df.columns) + 2, index=False)

        ws = writer.sheets[sheet]

        fmt_header = book.add_format({'bold': True, 'bg_color': '#E3F2FD', 'align': 'center'})
        fmt_ok   = book.add_format({'bg_color': '#E8F5E9'})
        fmt_ng   = book.add_format({'bg_color': '#FFEBEE'})
        fmt_grey = book.add_format({'bg_color': '#F5F5F5'})

        ncols_std = len(std_df.columns)
        ncols_sys = len(sys_df.columns)

        # 冻结首行
        ws.freeze_panes(1, 0)

        # 表头行
        ws.set_row(0, None, fmt_header)

        # 自动列宽
        def set_auto_width(df, offset_col):
            for i, col in enumerate(df.columns):
                max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).tolist()] + [8])
                ws.set_column(offset_col + i, offset_col + i, max_len + 2)

        set_auto_width(std_df, 0)
        set_auto_width(sys_df, ncols_std + 2)

        # 行着色
        def colorize(df, start_col):
            if '比对结果' not in df.columns:
                return
            for r, res in enumerate(df['比对结果'], start=1):
                if res == 'OK':
                    fmt = fmt_ok
                elif res == 'NG':
                    fmt = fmt_ng
                elif res == '未比对':
                    fmt = fmt_grey
                else:
                    fmt = None
                if fmt:
                    ws.set_row(r, None, fmt)

        colorize(std_df, 0)
        colorize(sys_df, ncols_std + 2)

        # 统计
        ok_std = int((std_df.get('比对结果') == 'OK').sum()) if '比对结果' in std_df.columns else 0
        ng_std = int((std_df.get('比对结果') == 'NG').sum()) if '比对结果' in std_df.columns else 0
        ok_sys = int((sys_df.get('比对结果') == 'OK').sum()) if '比对结果' in sys_df.columns else 0
        ng_sys = int((sys_df.get('比对结果') == 'NG').sum()) if '比对结果' in sys_df.columns else 0

        last_row = max(len(std_df), len(sys_df)) + 3
        ws.write(last_row,   0, f'导出时间: {datetime.now():%Y-%m-%d %H:%M:%S}')
        ws.write(last_row+1, 0, f'标准 OK: {ok_std} / NG: {ng_std}')
        ws.write(last_row+2, 0, f'系统 OK: {ok_sys} / NG: {ng_sys}')