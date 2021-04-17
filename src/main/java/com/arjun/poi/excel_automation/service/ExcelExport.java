package com.arjun.poi.excel_automation.service;


import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import com.arjun.poi.excel_automation.dao.RepositoryName;
import com.arjun.poi.excel_automation.repository.ITableNameRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static java.lang.String.format;
import static java.lang.String.valueOf;

@Slf4j
public class ExcelExport {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<RepositoryName> listOrderStatusHistory;

//    parameter should be your datatype<TableNameRepository> nameRepositorydatatype
//    public ExcelExport(<T>) {
//        this.listOrderStatusHistory = listOrderStatusHistory;
//        workbook = new XSSFWorkbook();
//    }


    private void writeHeaderLine() {
        sheet = workbook.createSheet("name goes here");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(12);
        style.setFont(font);

        createCell(row, 0, "coloumn name", style);

    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        }else  if (value instanceof Long) {
            cell.setCellValue((Long) value);
        }
        else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else if (value instanceof LocalDateTime) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yy HH:mm:ss");
            cell.setCellValue(((LocalDateTime) value).format(formatter));
        }
        else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }


    private void formatDate(XSSFSheet sheet) throws IOException {
        FormulaEvaluator evaluator1 = workbook.getCreationHelper().createFormulaEvaluator();

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum() ; rowIndex++ ){
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(8);
//            Adding custom formula
//            for end date based on status changed
//            =IF(A$2=A$3, (IF(MATCH(F$2, $E$1:$E$100 , 1),H$3,H$2)),  )
            cell.setCellFormula("IF(A$"+(rowIndex + 1)+"=A$"+(rowIndex + 2)+", (IF(MATCH(F$"+(rowIndex + 1)+", $E$"+(rowIndex+ 2)+":$E$"+(rowIndex + 10)+" , 1),H$"+(rowIndex + 2)+" ,H$"+(rowIndex + 1)+")), "+"\t"+")");
            evaluator1.evaluate(cell);
            log.info(row.getCell(8).getStringCellValue());
            Cell diffCell = row.getCell(9);
            diffCell.setCellFormula("INT(INT((I$"+(rowIndex + 1)+"-H$"+(rowIndex + 1)+")*1440)/60)&\":\"&INT((I$"+(rowIndex + 1)+"-H$"+(rowIndex + 1)+")*1440)&\":\"&IF(INT((I$"+(rowIndex + 1)+"-H$"+(rowIndex + 1)+")*1440*60)>60,INT((I$"+(rowIndex + 1)+"-H$"+(rowIndex + 1)+")*1440),INT((I$"+(rowIndex + 1)+"-H$"+(rowIndex + 1)+")*1440*60))");
            evaluator1.evaluate(diffCell);
//          for date time difference
// INT(INT((I2-H2)*1440)/60)&":"&INT((I2-H2)*1440)&":"&IF(INT((I2-H2)*1440*60)>60,INT((I2-H2)*1440),INT((I2-H2)*1440*60))
        }
        workbook.write(new FileOutputStream("test.xlsx"));
        workbook.close();
    }
    private void writeDataLines() {
        int rowCount = 1;

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

//        for (RepositoryName history : nameRepositioning) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;

            createCell(row, columnCount++, "", style);
            createCell(row, columnCount++, "", style);


        }
//    }

    public void export() throws IOException {
        this.writeHeaderLine();
        this.writeDataLines();
        workbook.write(new FileOutputStream("test.xlsx"));
        XSSFSheet  sheet =workbook.getSheetAt(0);
        this.formatDate(sheet);
        log.info(String.valueOf(sheet.getRow(1)));
        workbook.close();
    }
}
