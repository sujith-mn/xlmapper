package com.xlm;

import com.xlm.reader.SheetReader;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;


import java.io.File;


import static org.mockito.Mockito.mock;

public class SheetReaderTest {
    public   static void main(String args[]) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File("C:\\temp\\employee.xlsx"));
        SheetReader<Employee> sr=new SheetReader<Employee>(workbook);
        sr.setHeaderRow(0);
        Sheet datatypeSheet =   workbook.getSheetAt(0);
        System.out.println(sr.retrieveRows(datatypeSheet,Employee.class));
    }

    @Test
    public void shouldReturnList_whenExcelPassed() throws Exception {
        Workbook mockWorkbook = mock(Workbook.class);
        Sheet mockSheet = mock(Sheet.class);

        SheetReader<Employee> sr=new SheetReader<Employee>(mockWorkbook);

        Sheet datatypeSheet =   mockWorkbook.getSheetAt(0);
        System.out.println(sr.retrieveRows(datatypeSheet,Employee.class));

    }
}