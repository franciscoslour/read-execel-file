package com.example.excel.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.example.excel.dto.CellDetail;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {

    private static final Logger logger = LoggerFactory.getLogger(ExcelService.class);

    @EventListener
    public void readExcelFile(ApplicationReadyEvent event) {
        try {
            String fileLocation = System.getenv("EXCEL_FILE");
            FileInputStream fileInputStream = this.getFile(fileLocation);
            Sheet sheet = this.selectSheet(fileInputStream,0);

            for(int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++){
                List<CellDetail> rowValues = this.getRowValues(sheet, rowIndex);
                rowValues.forEach(row->{
                    logger.info(row.getColumnName() + " -> "+ row.getCellValue());
                });
            }

            fileInputStream.close();

        } catch (FileNotFoundException erro) {
            logger.error(erro.getMessage(), erro);
        } catch (IOException erro) {
            logger.error(erro.getMessage(), erro);
        }
    }

    public FileInputStream getFile(String fileLocation) throws FileNotFoundException {
        File file = new File(fileLocation);
        return new FileInputStream(file);
    }
    
    public Sheet selectSheet(FileInputStream fileInputStream, Integer sheetIndex) throws IOException{
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        return workbook.getSheetAt(sheetIndex);
    }

    public List<CellDetail> getRowValues(Sheet sheet, Integer rowIndex){
        List<CellDetail> rowValues = new ArrayList<>();
        try{
            for(Cell cell : sheet.getRow(rowIndex)){
                CellDetail cellDetail =  new CellDetail(sheet, cell, cell.getColumnIndex());
                rowValues.add(cellDetail);
            }
        }catch(NullPointerException error){
            logger.error(error.getMessage(), error);
        }
        return rowValues;
    }
}
