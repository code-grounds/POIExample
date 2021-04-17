package com.arjun.poi.excel_automation.service;

import com.arjun.poi.excel_automation.dao.RepositoryName;
import com.arjun.poi.excel_automation.repository.ITableNameRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

import static org.apache.poi.ss.util.CellUtil.createCell;

@Slf4j
@Service
public class ServiceClass {

    private XSSFWorkbook FlowOrderStatusHistory;
    private XSSFSheet orderStatusHistorySheet;
    private List<RepositoryName> orderStatusHistoryList;

    @Autowired
    ITableNameRepository tableNameRepository;


    public void generateExcel() throws IOException {

//        List<RepoName> variable = tableNameRepository.findByMethodName(value);

//        ExcelExport excelExport = new ExcelExport(variable);

//        excelExport.export();
        log.info("get repo {}");
    }

}
