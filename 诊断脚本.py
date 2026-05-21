#!/usr/bin/env python3
"""
终极诊断脚本 —— 绕过浏览器，直接检查服务端返回的文件名
"""
import io
import os
import sys
import pandas as pd
from app import create_app

def check_download_name():
    app = create_app()
    app.config['TESTING'] = True

    with app.test_client() as client:
        # 构造一个简单 Excel
        df = pd.DataFrame({'姓名': ['张三', '李四'], '金额': [100, 200]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, '测试文件.xlsx'),
            'key_cols': '姓名',
            'calc_cols': '金额',
            'agg_func_default': 'sum',
            'sheets_mode': '1'
        }

        resp = client.post('/subtotal/process', data=data,
                           content_type='multipart/form-data')

        print(f"状态码: {resp.status_code}")
        if resp.status_code != 200:
            print(f"错误信息: {resp.get_json()}")
            return

        content_disp = resp.headers.get('Content-Disposition', '')
        print(f"Content-Disposition: {content_disp}")

        if 'subtotal_测试文件_小计.xlsx' in content_disp:
            print("\n✅ 服务端返回的文件名正确！问题在浏览器缓存。")
            print("请尝试：使用隐私模式下载，或按 Ctrl+Shift+Del 清除缓存。")
        else:
            print("\n❌ 服务端返回的文件名不正确。")
            print("可能原因：引擎或蓝图未真正生效。请检查是否还有残留的旧进程。")

if __name__ == '__main__':
    check_download_name()