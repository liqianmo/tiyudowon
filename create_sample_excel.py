#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例 Excel 文件
"""


import pandas as pd

def create_sample_excel():
    """创建示例 Excel 文件"""
    data = {
        '编号': [
            'NO001',
            'NO002', 
            'NO003',
            'NO004',
            'NO005'
        ],
        '姓名': [
            '王晓燕',
            '李明',
            '张华',
            '刘红',
            '陈强'
        ],
        '性别': [
            '女',
            '男',
            '男',
            '女',
            '男'
        ],
        '年龄': [
            25,
            32,
            28,
            30,
            35
        ],
        '俱乐部名称': [
            '天天俱乐部',
            '阳光体育',
            '健康运动会',
            '天天俱乐部',
            '阳光体育'
        ],
        '部门': [
            '研发部',
            '市场部',
            '财务部',
            '人事部',
            '技术部'
        ],
        '城市': [
            '北京',
            '上海',
            '广州',
            '深圳',
            '杭州'
        ],
        '照片链接': [
            'https://picsum.photos/300/400?random=1',
            'https://picsum.photos/300/400?random=2',
            'https://picsum.photos/300/400?random=3',
            'https://picsum.photos/300/400?random=4',
            'https://picsum.photos/300/400?random=5'
        ],
        '备注': [
            '优秀员工',
            '团队领导',
            '技术专家',
            '新人培训',
            '资深员工'
        ]
    }
    
    df = pd.DataFrame(data)
    df.to_excel('示例数据.xlsx', index=False)
    print("示例 Excel 文件已创建: 示例数据.xlsx")

if __name__ == "__main__":
    create_sample_excel()
