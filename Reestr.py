#!/usr/bin/env python
# -*- coding: utf8 -*-
# Reestr App

import os
import sys
import csv
import xlrd

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet
from concurrent.futures import ThreadPoolExecutor
from PyQt6.QtWidgets import (
    QApplication,
    QGridLayout,
    QPushButton,
    QLineEdit,
    QCheckBox,
    QDialog,
    QMainWindow,
    QWidget,
    QFileDialog,
    QDialogButtonBox,
)

class ReestrWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reestr")
        self.generalLayout = QGridLayout()
        centralWidget = QWidget(self)
        centralWidget.setLayout(self.generalLayout)
        self.setCentralWidget(centralWidget)
        
        if os.path.exists('settings.txt') == False:
            open(r"settings.txt", "w").writelines(["Input files path\n","Output files path\n","Person\n","Organization\n","sign.png\n","False\n"])
        elif len(open(r"settings.txt", "r").readlines()) != 6:
            open(r"settings.txt", "w+").writelines(["Input files path\n","Output files path\n","Person\n","Organization\n","sign.png\n","False\n"])
        lines = open(r"settings.txt", "r").readlines()

        self.inputPath = QLineEdit(lines[0].strip('\n'))
        self.outputPath = QLineEdit(lines[1].strip('\n'))
        chooseInputPath = QPushButton("Choose Input Path")
        chooseOutputPath = QPushButton("Choose Output Path")
        personButton = QPushButton("Person")
        organizationButton = QPushButton("Organization")
        signButton = QPushButton("Sign")
        deleteCheckbox = QCheckBox("Delete source files after convertation")
        convertButton = QPushButton("Convert")
        exitButton = QPushButton("Exit")

        boolean = False
        if lines[5].strip('\n') == 'True':
            boolean = True
        deleteCheckbox.setChecked(boolean)
        
        self.generalLayout.addWidget(self.inputPath, 0, 0, 1, 3)
        self.generalLayout.addWidget(chooseInputPath, 0, 3)
        self.generalLayout.addWidget(self.outputPath, 1, 0, 1, 3)
        self.generalLayout.addWidget(chooseOutputPath, 1, 3)
        self.generalLayout.addWidget(personButton, 2, 0)
        self.generalLayout.addWidget(organizationButton, 2, 1)
        self.generalLayout.addWidget(signButton, 2, 2)
        self.generalLayout.addWidget(deleteCheckbox, 3, 0, 1, 4)
        self.generalLayout.addWidget(convertButton, 4, 2)
        self.generalLayout.addWidget(exitButton, 4, 3)

        chooseInputPath.clicked.connect(lambda: self.setPath(0))
        chooseOutputPath.clicked.connect(lambda: self.setPath(1))
        signButton.clicked.connect(self.setSignPath)
        personButton.clicked.connect(self.openPersonDialog)
        organizationButton.clicked.connect(self.openOrganizationDialog)
        deleteCheckbox.toggled.connect(lambda: self.setLine(5, str(deleteCheckbox.isChecked())))
        convertButton.clicked.connect(lambda: Convertor.convert(self.inputPath.text(), self.outputPath.text(), deleteCheckbox.isChecked()))
        exitButton.clicked.connect(self.close)

    def setPath(self, lineNum):
        inOut = [self.inputPath, self.outputPath]
        path = QFileDialog.getExistingDirectory(self, "Select Folder")

        if path != "":
            inOut[lineNum].setText(path + '/')
            lines = open(r"settings.txt", "r").readlines()
            lines[lineNum] = path + '/' + '\n'
            open(r"settings.txt", "w").writelines(lines)

    def setSignPath(self):
        path = QFileDialog.getOpenFileName(self, "Select a File")

        if path[0] != "":
            lines = open(r"settings.txt", "r").readlines()
            lines[4] = path[0] + '\n'
            open(r"settings.txt", "w").writelines(lines)

    def openPersonDialog(self):
        person = PersonDialog(self)
        person.exec()

    def openOrganizationDialog(self):
        organization = OrganizationDialog(self)
        organization.exec()

    @staticmethod
    def getLine(lineNum):
        lines = open(r"settings.txt", "r").readlines()
        line = lines[lineNum]
        return line

    def setLine(self, lineNum, data):
        lines = open(r"settings.txt", "r").readlines()
        lines[lineNum] = data + '\n'
        open(r"settings.txt", "w").writelines(lines)

class PersonDialog(QDialog):
    def __init__(self, parent=ReestrWindow):
        super().__init__(parent)
        self.setWindowTitle("Person")
        self.layout = QGridLayout()
        personEdit = QLineEdit(parent.getLine(2).strip('\n'))
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.layout.addWidget(personEdit, 0, 0, 1, 3)
        self.layout.addWidget(self.buttonBox, 1, 1, 1, 2)
        self.setLayout(self.layout)
        self.accepted.connect(lambda: parent.setLine(2, personEdit.text()))

class OrganizationDialog(QDialog):
    def __init__(self, parent=ReestrWindow):
        super().__init__(parent)
        self.setWindowTitle("Organization")
        self.layout = QGridLayout()
        orgEdit = QLineEdit(parent.getLine(3).strip('\n'))
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.layout.addWidget(orgEdit, 0, 0, 1, 3)
        self.layout.addWidget(self.buttonBox, 1, 1, 1, 2)
        self.setLayout(self.layout)
        self.accepted.connect(lambda: parent.setLine(3, orgEdit.text()))

class Convertor:
    sellerTitle = ["default", "Вента", "БАДМ", "Оптіма", "Юніфарма", "Дельта", "Фіто-Лек"]
    
    @staticmethod
    def convert(inputpath, outputpath, deleteCheckbox):
        inputFiles = os.listdir(inputpath)

        filteredFilenames = [os.path.join(inputpath, file) for file in inputFiles if file.endswith((".xls", ".csv"))]

        outputFilenames = [os.path.join(outputpath, os.path.basename(file)) for file in filteredFilenames]

        preprocessedFiles = []
        for file in filteredFilenames:
            if file.endswith(".csv"):
                preprocessedFiles.append(Convertor.convertCsvToArray(file))
            else:
                preprocessedFiles.append(Convertor.readWorkbook(file))

        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(Convertor.process, preprocessedFiles[i], outputFilenames[i]) for i in range(len(preprocessedFiles))]
            for future in futures:
                future.result()

        if deleteCheckbox == True:
            for f in filteredFilenames:
                if os.path.isfile(f):
                    os.remove(f)
        
    def readWorkbook(file):
        try:
            workBook = xlrd.open_workbook(file, formatting_info=True)
            return workBook
        except Exception:
            return None
        
    def process(preprocessedFile, outputFilename):
        fileToProcess = preprocessedFile

        if "Delta" in outputFilename:
            Convertor.processDelta(preprocessedFile, outputFilename)
        else:
            try:
                dateOfDocument = None
                reestrType = 0
                
                if isinstance(fileToProcess, list):

                    if "БаДМ" in fileToProcess[0][1]:
                        dateCell = fileToProcess[0][2]
                        predate = dateCell.rfind(' ')
                        dateOfDocument = dateCell[predate + 1:]
                        reestrType = 2
                        processedFile = fileToProcess
                    elif "Юніфарма" in fileToProcess[0][1]:
                        dateCell = fileToProcess[0][2]
                        predate = dateCell.rfind(' ')
                        dateOfDocument = dateCell[predate + 1:]
                        reestrType = 4
                        processedFile = fileToProcess
                else:
                    wbSheetToProcess = fileToProcess.sheet_by_index(0)

                    if "ВЕНТА. ЛТД" in wbSheetToProcess.cell(4, 1).value:
                        dateCell = wbSheetToProcess.cell(4, 2).value
                        predate = dateCell.rfind(' ')
                        dateOfDocument = dateCell[predate + 1:]
                        reestrType = 1
                        processedFile = Convertor.convertXlsToArray(wbSheetToProcess, 4, 2)
                    elif "Оптiма-Фарм" in wbSheetToProcess.cell(8, 1).value:
                        dateCell = wbSheetToProcess.cell(9, 2).value
                        predate = dateCell.rfind(' ')
                        dateOfDocument = dateCell[predate + 1:]
                        reestrType = 3
                        processedFile = Convertor.convertXlsToArray(wbSheetToProcess, 8, 3)
                outputFilename = Convertor.mkdirs(Convertor.sellerTitle[reestrType], outputFilename, dateOfDocument)
                Convertor.convertToPdf(processedFile, outputFilename, dateOfDocument, reestrType)
            except Exception:
                print("Не вдалось сконвертувати: " + outputFilename)

    def processDelta(preprocessedFile, outputFilename):
        try:
            dateOfDocument = None
            reestrType = 5
            preprocessedWbSheet = preprocessedFile.sheet_by_index(0)
            deltaRow = '"ДЕЛЬТА МЕДІКЕЛ" ліцензія'
            firstRow = 0

            for row in range(0, preprocessedWbSheet.nrows - 1):
                 if deltaRow in preprocessedWbSheet.cell_value(row, 2):
                     firstRow = row
                     break

            rowCount = 0
            for row in range(0, preprocessedWbSheet.nrows - 1):
                 if deltaRow in preprocessedWbSheet.cell_value(row, 2):
                     rowCount += 1

            dateCell = preprocessedWbSheet.cell_value(firstRow, 8)
            predate = dateCell.rfind(' ')
            dateOfDocument = dateCell[predate + 1:]
            processedFile = [[0 for x in range(preprocessedWbSheet.ncols)] for y in range(preprocessedWbSheet.nrows)]

            if firstRow == 11:
                for i in range(firstRow, firstRow + rowCount):
                    for cellIn in range(0, preprocessedWbSheet.row_len(i) - 1):
                            processedFile[i - 2][cellIn] = preprocessedWbSheet.cell_value(i, cellIn)
                    processedFile[i - 2][22] = "Відповідає"
            else:
                for i in range(firstRow, firstRow + rowCount // 2):
                    for cellIn in range(0, preprocessedWbSheet.row_len(i) - 1):
                        processedFile[i - 18 + rowCount // 2][cellIn] = preprocessedWbSheet.cell_value(i, cellIn)
            processedFile = [[y for y in x if y != 0 and y != ''] for x in processedFile]
            processedFile = [y for y in processedFile if y != []]
            outputFilename = Convertor.mkdirs(Convertor.sellerTitle[reestrType], outputFilename, dateOfDocument)
            Convertor.convertToPdf(processedFile, outputFilename, dateOfDocument, reestrType)
        except Exception:
            print("Невідомий формат")
            
    def convertCsvToArray(csvFile):
        preprocessedFile = []
    
        with open(csvFile, 'r', encoding='Cp1251') as file:
            csvReader = csv.reader(file, dialect='excel')
            rows = list(csvReader)
            csvType = rows[0]

            if 'Додаток' in csvType[0]:
                rowNum = 0
                if len(rows) < 8:
                    for cell in rows[1:]:
                        tempRow = []
                        tempRow.append(cell[0])
                else:
                    for cell in rows[8:]:
                        tempRow = []
                        rowNum += 1
                        tempRow.append(str(rowNum))
                        tempRow.extend(cell[:8])
                        tempRow.append('Відповідає')
                preprocessedFile.append(tempRow)
            elif 'Реєстр' in csvType[0]:
                for cell in rows[3:]:
                    tempRow = cell[:9]
                    tempRow.append('Відповідає')
                    preprocessedFile.append(tempRow)
        return preprocessedFile
    
    def convertToPdf(processedFile, outputFilename, dateOfDocument, reestr_type):
        doc = SimpleDocTemplate(outputFilename.replace(".xls", ".pdf").replace(".csv", ".pdf"), pagesize=landscape(A4))
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        pdfmetrics.registerFontFamily('Arial',normal='Arial',bold='Arial',italic='Arial',boldItalic='Arial')
        elements = []
    
        table_data = []
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])

        commonCellStyle = getSampleStyleSheet()["BodyText"]
        commonCellStyle.alignment = TA_CENTER
        commonCellStyle.fontName = 'Arial'
        commonCellStyle.fontSize = 10

        leftCellStyle = getSampleStyleSheet()["BodyText"]
        leftCellStyle.alignment = TA_LEFT
        leftCellStyle.fontName = 'Arial'
        leftCellStyle.fontSize = 10

        header_text = "Реєстр<br />лікарських засобів, які надійшли до суб'єкта господарювання<br />" + ReestrWindow.getLine(3)
        header_cell = Paragraph(header_text, leftCellStyle)

        columns = ["№ з/п", "Назва постачальника та номер ліцензії", "Номер та дата накладної",
                   "Назва лікарського засобу та його лікарська форма, дата реєстрації та номер реєстраційного посвідчення",
                   "Назва виробника", "Номер серії", "Номер і дата сертифіката якості виробника", "Кількість одержаних упаковок",
                   "Термін придатності лікарського засобу", "Результат контролю уповноваженою особою"]
        columns_text = []
        for column in columns:
            column_cell = Paragraph(column, commonCellStyle)
            columns_text.append(column_cell)
        table_data.append(columns_text)
    
        for row in processedFile:
            row_cells = []
            for cell in row:
                cell_cell = Paragraph(str(cell), commonCellStyle)
                row_cells.append(cell_cell)
            table_data.append(row_cells)
    
        footer_text = "Результат вхідного контролю якості лікарських засобів здійснив — уповноважена особа " + ReestrWindow.getLine(2) + '<br />' + dateOfDocument
        footer_cell = Paragraph(footer_text, leftCellStyle)
        footer_cell.keepWithNext = True

        table = Table(table_data, colWidths=[0.4*inch, 1.2*inch, 0.8*inch, 2.2*inch, 1*inch, 0.8*inch, 0.9*inch, 0.8*inch, 0.9*inch, 0.9*inch])
        table.setStyle(table_style)
        elements.append(header_cell)
        elements.append(table)
        elements.append(footer_cell)

        scan_path = ReestrWindow.getLine(4).strip('\n')
        if not os.path.exists(scan_path):
            print("Не вибрано скан штампа.")
        else:
            image_cell = Image(scan_path)
            elements.append(image_cell)
    
        doc.build(elements)
        
    def mkdirs(sellerTitle, outputFilenames, dateOfDocument):
        month = dateOfDocument[dateOfDocument.index(".") + 1:dateOfDocument.rindex(".")]
        months = {"01": "Січень", "02": "Лютий", "03": "Березень", "04": "Квітень", "05": "Травень",
                  "06": "Червень", "07": "Липень", "08": "Серпень", "09": "Вересень", "10": "Жовтень",
                  "11": "Листопад", "12": "Грудень"}
        month = months.get(month, "")
        outputFilenames = outputFilenames.replace(outputFilenames[outputFilenames.rindex('/'):],
                                '/' + sellerTitle + '/' + dateOfDocument[6:10] +
                                '/' + month + '/' + dateOfDocument +
                                outputFilenames[outputFilenames.rindex('/'):])
        directory = outputFilenames[:outputFilenames.rindex('/')]
        os.makedirs(directory, exist_ok=True)
        return outputFilenames
    
    def convertXlsToArray(wbSheetToProcess, first_row_offset, last_row_offset):
        processedFile = [[0 for x in range(9)] for y in range(wbSheetToProcess.nrows - first_row_offset - last_row_offset)] 

        for row in range(first_row_offset, wbSheetToProcess.nrows - last_row_offset):
            for cell in range(9):
                processedFile[row - first_row_offset][cell] = wbSheetToProcess.cell_value(row, cell)
            processedFile[row - first_row_offset].append("Відповідає")
        return processedFile

def main():
    app = QApplication([])
    window = ReestrWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
