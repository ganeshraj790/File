package com.example.file.dataservice;

import com.example.file.entity.Invoice;
import com.example.file.repository.InvoiceRepo;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class DataServiceImpl implements DataService{


    @Value("${app.upload.file:${user.home}}")
    public String EXCEL_FILE_PATH;

    @Autowired
    InvoiceRepo invoiceRepo;

    Workbook workbook;

    public List<Invoice> getExcelDataAsList() {

        List<String> list = new ArrayList<String>();

        DataFormatter dataFormatter = new DataFormatter();

        try {
            workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));
        } catch (EncryptedDocumentException | IOException e) {
            e.printStackTrace();
        }
        System.out.println("-------Workbook has '" + workbook.getNumberOfSheets() + "' Sheets-----");

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Getting number of columns in the Sheet
        int noOfColumns = sheet.getRow(0).getLastCellNum();
        System.out.println("-------Sheet has '"+noOfColumns+"' columns------");

        // Using for-each loop to iterate over the rows and columns
        for (Row row : sheet) {
            for (Cell cell : row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                list.add(cellValue);
            }
        }

        List<Invoice> invList = createList(list, noOfColumns);

        // Closing the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return invList;
    }

    private List<Invoice> createList(List<String> excelData, int noOfColumns) {

        ArrayList<Invoice> invList = new ArrayList<Invoice>();

        int i = noOfColumns;
        do {
            Invoice inv = new Invoice();

            inv.setName(excelData.get(i));
            inv.setAmount(Double.valueOf(excelData.get(i + 1)));
            inv.setNumber(excelData.get(i + 2));
            inv.setReceivedDate(excelData.get(i + 3));

            invList.add(inv);
            i = i + (noOfColumns);

        } while (i < excelData.size());
        return invList;
    }
    @Override
    public int saveExcelData(List<Invoice> invoices) {
        invoices = invoiceRepo.saveAll(invoices);
        return invoices.size();
    }

}
