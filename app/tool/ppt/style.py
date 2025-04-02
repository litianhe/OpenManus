from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Tuple


class PPTStyle(Enum):
    """
    PPT style enumeration
    """
    FRESH_GREEN = "FRESH_GREEN"  # Simple Fresh Green
    BUSINESS_BLUE = "BUSINESS_BLUE"  # Business Fresh Blue

@dataclass
class StyleConfig:
    """ PPT style configuration data class"""
    theme_color: Dict[str, Tuple[int, int, int]]
    fonts: Dict[str, Dict[str, Any]]
    layout: Dict[str, int]
    effects: Dict[str, float]

# 预定义的风格配置 / Predefined style configurations
STYLE_CONFIGS = {
    PPTStyle.FRESH_GREEN: StyleConfig(
        theme_color={
            "primary": (67, 134, 84),      # 深绿色作为强调色
            "secondary": (228, 241, 225),  # 极浅绿作为渐变起始色
            "accent": (164, 203, 164),     # 中绿色作为渐变结束色
            "background": (255, 255, 255), # 纯白作为基础背景
            "text_dark": (33, 33, 33),     # 近黑色文字增加可读性
            "text_light": (89, 89, 89),    # 灰色文字用于次要信息
            "highlight": (255, 200, 87)    # 金色作为强调
        },
        fonts={
            "title": {
                "name": "Microsoft YaHei",  # 微软雅黑
                "size": 36,
                "bold": True
            },
            "subtitle": {
                "name": "Microsoft YaHei Light",
                "size": 28,
                "bold": False
            },
            "body": {
                "name": "Microsoft YaHei",
                "size": 18,
                "bold": False
            },
            "data": {
                "name": "Microsoft YaHei",
                "size": 16,
                "bold": False
            }
        },
        layout={
            "page_margin": 80,       # 页边距
            "header_height": 120,    # 页眉
            "footer_height": 30,     # 页脚高度
            "grid_gutter": 30        # 网格间距
        },
        effects={
            "shape_opacity": 0.12,   # 形状透明度
            "line_weight": 1.2       # 线条粗细
        }
    ),
    PPTStyle.BUSINESS_BLUE: StyleConfig(
        theme_color={
            "primary": (51, 122, 183),     # 商务蓝 / Business blue
            "secondary": (176, 196, 222),  # 淡雅蓝 / Light elegant blue
            "accent": (30, 73, 110),       # 深邃蓝 / Deep blue
            "background": (248, 251, 255), # 清爽白 / Fresh white
            "text_dark": (44, 62, 80),     # 深色文字 / Dark text
            "text_light": (149, 165, 166), # 浅色文字 / Light text
            "highlight": (255, 168, 46)    # 点缀橙 / Accent orange
        },
        fonts={
            "title": {"name": "思源黑体", "size": 32, "bold": True},
            "subtitle": {"name": "思源黑体", "size": 24},
            "body": {"name": "微软雅黑", "size": 18},
            "data": {"name": "Segoe UI", "size": 16}
        },
        layout={
            "page_margin": 80,      # 页面边距 / Page margin
            "header_height": 120,   # 页眉高度 / Header height
            "footer_height": 70,    # 页脚高度 / Footer height
            "grid_gutter": 30       # 栅格间距 / Grid gutter
        },
        effects={
            "shape_opacity": 0.12,  # 形状透明度 / Shape opacity
            "line_weight": 1.8      # 装饰线粗细 / Line weight
        }
    )
}
