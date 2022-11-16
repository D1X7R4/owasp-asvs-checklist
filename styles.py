from enum import Enum

class Styles(Enum):

    conditional_format = {'font_name': 'Avenir','font_size': 11,'valign': 'vcenter', 'align': 'center','border': 5, 'border_color': 'white' }
    pass_format = {'bg_color':'#B6D7A8', 'font_color': '#38761D' }
    fail_format = {'bg_color':'#FFC7CE', 'font_color': '#9C0006'}
    na_format = {'bg_color':'#CCCCCC', 'font_color': '#666666' }
    testing_format = {'bg_color':'#F79646', 'font_color': '#E26B0A'}

    title_format = {'font_name': 'Avenir', 'font_size': 24, 'bold': True, 'text_wrap': True, 'valign': 'vcenter'}
    subtitle_format = {'font_name': 'Avenir', 'font_size': 18, 'text_wrap': True, 'valign': 'vcenter'}
    version_format = {'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'valign': 'vcenter'}
    
    title_section_format = {'bold': True, 'font_name': 'Avenir', 'font_size': 15, 'text_wrap': True, 'font_color': '#499FFF', 'valign': 'vcenter', 'bottom': 5, 'bottom_color': '#499FFF'}
    fields_format = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#C0C0C0', 'valign': 'vcenter', 'align': 'center'}
    id_format = {'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#000000', 'valign': 'vcenter','align': 'center'}
    requirements_format = {'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#000000', 'valign': 'vcenter'}
    
    # Conditional levels
    level1_format = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#33CCCC', 'bg_color': '#33CCCC', 'border': 5, 'border_color': 'white', 'valign': 'vcenter'}
    level2_format = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#99CC00', 'bg_color': '#99CC00', 'border': 5, 'border_color': 'white','valign': 'vcenter'}
    level3_format = {'font_name': 'Avenir', 'font_size': 11, 'text_wrap': True, 'font_color': '#FF9900', 'bg_color': '#FF9900','border': 5, 'border_color': 'white', 'valign': 'vcenter'}
    
    statistics_titles = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#23548D'}
    statistics_section_title = {'bold': True, 'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'bg_color': '#23548D'}
    statistics_section_data = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'bg_color': '#347ED4'}
    statistics_data = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#347ED4'}
    statistics_percentage = {'font_name': 'Avenir', 'font_size': 11, 'font_color': 'white', 'valign': 'vcenter', 'align': 'center', 'bg_color': '#347ED4', 'num_format': 10}