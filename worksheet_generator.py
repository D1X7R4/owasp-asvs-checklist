from xlsxwriter.workbook import Workbook 
from styles import Styles

class WorkSheetGenerator:
    
    def __init__(self, workbook: Workbook, worksheet_title, worksheet_name,worksheet_shortname, name, version):
        self.workbook = workbook
        self.worksheet_title = worksheet_title
        self.worksheet_name = worksheet_name
        self.worksheet_shortname = worksheet_shortname
        self.name = name
        self.version = version
        self.worksheet = workbook.add_worksheet(worksheet_shortname)

    def generateWorksheet(self, requirements):
        self.worksheet.hide_gridlines(2)
        self.worksheet.insert_image('B2', 'assets/logo.png', {'object_positon': 2,'x_scale': 0.41, 'y_scale': 0.30})
        self.worksheet.set_column("B:B", 15)
        self.worksheet.set_column("C:C", 100)
        self.worksheet.set_column("D:F", 4.17)
        #self.worksheet.set_column("G:G", 10)
        self.worksheet.write("C3", self.worksheet_title,self.workbook.add_format(Styles.title_format.value))
        self.worksheet.write("C4", self.worksheet_name,self.workbook.add_format(Styles.subtitle_format.value))
        self.worksheet.write("C5", f'{self.name} {self.version}', self.workbook.add_format(Styles.version_format.value))
        row = 8
        for section in requirements:
            row = self.writeRequirements(section['Name'], section['Items'], row)

    def writeRequirements(self, sectionName, requirements, row):
        self.worksheet.merge_range(f"B{row}:G{row}", sectionName, self.workbook.add_format(Styles.title_section_format.value))
        row += 2
        self.worksheet.write(f"B{row}", "ID", self.workbook.add_format(Styles.fields_format.value))
        self.worksheet.write(f"C{row}", "Detailed Verification Requirement", self.workbook.add_format(Styles.fields_format.value))
        self.worksheet.write(f"D{row}", "L1", self.workbook.add_format(Styles.fields_format.value))
        self.worksheet.write(f"E{row}", "L2", self.workbook.add_format(Styles.fields_format.value))
        self.worksheet.write(f"F{row}", "L3", self.workbook.add_format(Styles.fields_format.value))
        # self.worksheet.write(f"G{row}", "Test Case", self.workbook.add_format(Styles.fields_format.value))
        self.worksheet.write(f"G{row}", "Status", self.workbook.add_format(Styles.fields_format.value))
        row += 2
        for req in requirements:
            self.worksheet.set_row(row-1, 55)
            self.worksheet.write(f"B{row}", req['Shortcode'], self.workbook.add_format(Styles.id_format.value))
            self.worksheet.write(f"C{row}", req['Description'], self.workbook.add_format(Styles.requirements_format.value))
            self.worksheet.write(f"D{row}", req['L1']['Requirement'], self.workbook.add_format(Styles.level1_format.value) if req['L1']['Required'] else self.workbook.add_format(Styles.requirements_format.value))
            self.worksheet.write(f"E{row}", req['L2']['Requirement'], self.workbook.add_format(Styles.level2_format.value) if req['L2']['Required'] else self.workbook.add_format(Styles.requirements_format.value))
            self.worksheet.write(f"F{row}", req['L3']['Requirement'], self.workbook.add_format(Styles.level3_format.value) if req['L3']['Required'] else self.workbook.add_format(Styles.requirements_format.value))
            self.worksheet.data_validation(f"G{row}", {'validate': 'list', 'source': ['Pass', 'Fail', 'N/A', 'TBT']})
            self.worksheet.write(f"G{row}", None, self.workbook.add_format(Styles.conditional_format.value))
            self.worksheet.conditional_format(f"G{row}", {'type': 'text', 'criteria': 'containing', 'value': 'Pass', 'format': self.workbook.add_format(Styles.pass_format.value) })
            self.worksheet.conditional_format(f"G{row}", {'type': 'text', 'criteria': 'containing', 'value': 'N/A', 'format': self.workbook.add_format(Styles.na_format.value)})
            self.worksheet.conditional_format(f"G{row}", {'type': 'text', 'criteria': 'containing', 'value': 'TBT', 'format': self.workbook.add_format(Styles.testing_format.value)})
            self.worksheet.conditional_format(f"G{row}", {'type': 'text', 'criteria': 'containing', 'value': 'Fail', 'format': self.workbook.add_format(Styles.fail_format.value)})

            row+=1
        row+=1
        return row