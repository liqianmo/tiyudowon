#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例 Excel 文件
"""

import pandas as pd

def create_sample_excel():
    """创建示例 Excel 文件"""
    data = {
        '姓名': [
            '王晓燕',
            '李明',
            '张华',
            '刘红',
            '陈强'
        ],
        '俱乐部名称': [
            '天天俱乐部',
            '阳光体育',
            '健康运动会',
            '天天俱乐部',
            '阳光体育'
        ],
        '照片链接': [
            'https://picsum.photos/300/400?random=1',
            'https://picsum.photos/300/400?random=2',
            'https://picsum.photos/300/400?random=3',
            'https://picsum.photos/300/400?random=4',
            'https://picsum.photos/300/400?random=5'
        ]
    }
    
    df = pd.DataFrame(data)
    df.to_excel('示例数据.xlsx', index=False)
    print("示例 Excel 文件已创建: 示例数据.xlsx")

if __name__ == "__main__":
    create_sample_excel()
