from xlsxwriter.workbook import Workbook 
from overview_styles import OverviewStyles

class OverviewWorksheet:
    
    def __init__(self, workbook: Workbook, worksheet_title, categories):
        self.workbook = workbook
        self.worksheet_title = worksheet_title
        self.categories = categories
        self.worksheet = self.workbook.add_worksheet(worksheet_title)
        self.generateOverviewWorksheet()

    def generateChart(self):
        chart = self.workbook.add_chart({'type': 'radar', 'subtype': 'filled'})
        chart.add_series({'name_font': {'color':'#595959'},'name':'Percentage (total/completed)', 'categories': f'={self.worksheet_title}!$A$2:$A$15', 'values': f'={self.worksheet_title}!$J$2:$J$15','line': {'color': '#4F81BD', "width": 4, 'transparency': 70}, 'marker': {'type': 'circle', 'size': 4}, "fill": {"none": True}})
        chart.set_y_axis(
            {'min': 0.0, 'max': 1.0, 
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': '#D9D9D9', 'width': 0.75}
                },
                'line': {'none': True},
                'num_font':{'color': '#595959'}
            }
        )
        chart.set_chartarea({'name_font':{'color': '#595959'}})
        chart.set_x_axis(
            {'num_font':{'color': '#595959'}}
        )
        chart.set_size({'width': 1458, 'height': 415 })
        chart.set_legend({'position': 'bottom', 'font':{'color': '#595959'}})
        chart.set_title({'name': 'Passed requirements', 'name_font':{'color': '#595959'}})
        self.worksheet.insert_chart('A16', chart)

    def generateOverviewWorksheet(self):
        self.worksheet.set_column("A:A", 50)
        self.worksheet.set_column("B:K", 15)

        self.worksheet.write("A1", "Section", self.workbook.add_format(OverviewStyles.TABLE_CATEGORY_TITLE.value))
        self.worksheet.write("B1", "L1", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("C1", "L2", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("D1", "L3", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("E1", "Total", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("F1", "L1 (Pass)", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("G1", "L2 (Pass)", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("H1", "L3 (Pass)", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("I1", "Total (Pass)", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("J1", "Pecentage", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        self.worksheet.write("K1", "Current Level", self.workbook.add_format(OverviewStyles.TABLE_TITLES.value))
        
        for index in range(2,len(self.categories)+1):
            self.worksheet.write(f"A{index}", self.categories[index-2][0], self.workbook.add_format(OverviewStyles.TABLE_CATEGORIES.value))
            self.worksheet.write_formula(f"B{index}", f'=COUNTIFS({self.categories[index-2][1]}!D:D, "✓", {self.categories[index-2][1]}!G:G, "<>N/A")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"C{index}", f'=COUNTIFS({self.categories[index-2][1]}!E:E, "✓", {self.categories[index-2][1]}!G:G, "<>N/A")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"D{index}", f'=COUNTIFS({self.categories[index-2][1]}!F:F, "✓", {self.categories[index-2][1]}!G:G, "<>N/A")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"E{index}", f'=MAX(B{index}:D{index})', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"F{index}", f'=COUNTIFS({self.categories[index-2][1]}!D:D, "✓", {self.categories[index-2][1]}!G:G, "Pass")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"G{index}", f'=COUNTIFS({self.categories[index-2][1]}!E:E, "✓", {self.categories[index-2][1]}!G:G, "Pass")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"H{index}", f'=COUNTIFS({self.categories[index-2][1]}!F:F, "✓", {self.categories[index-2][1]}!G:G, "Pass")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"I{index}", f'=MAX(F{index}:H{index})', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))
            self.worksheet.write_formula(f"J{index}", f'=I{index}/E{index}', self.workbook.add_format(OverviewStyles.TABLE_DATA_PERCENTAGE.value))
            self.worksheet.write_formula(f"K{index}", f'=_xlfn.IFS(AND(D{index}<>0,D{index}=H{index}),"L3",AND(C{index}<>0,C{index}=G{index}),"L2",AND(B{index}<>0,B{index}=F{index}), "L1", TRUE, "")', self.workbook.add_format(OverviewStyles.TABLE_DATA.value))

        self.generateChart()