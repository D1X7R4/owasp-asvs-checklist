from enum import Enum

class OverviewStyles(Enum):
    TABLE_TITLES = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#23548D'}
    TABLE_CATEGORY_TITLE = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'bg_color': '#23548D'}
    TABLE_CATEGORIES = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'bg_color': '#23548D'}
    TABLE_DATA = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#347ED4'}
    TABLE_DATA_PERCENTAGE = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#347ED4', 'num_format': 10}
