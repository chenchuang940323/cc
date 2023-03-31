package com.cc.cctest.test;

import com.cc.cctest.pojo.PersonMessage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import util.ExcelUtil;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Test1 {

    public static void main(String[] args) {

        String filePath = "C:\\Users\\科大国创\\Desktop\\test3.xlsx";

        getExcel(filePath);

    }

    public static List<PersonMessage> getExcel(String filePath){

        List<PersonMessage> list = new ArrayList<>();





        try {
            Workbook workbook = ExcelUtil.getWorkbook(new File(filePath));




            Sheet sheet = workbook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum();

            for(int i = 1; i <= rowNum; i++){
                PersonMessage personMessage = new PersonMessage();

                String name = ExcelUtil.getCellStringValue(sheet, i, 0);
                personMessage.setName(name);

                String max = ExcelUtil.getCellStringValue(sheet, i, 1);
                personMessage.setMax(max);

                double high = ExcelUtil.getCellNumericValue(sheet, i, 2);
                personMessage.setHigh(high);

                double weight = ExcelUtil.getCellNumericValue(sheet, i, 3);
                personMessage.setWeight(weight);

                String date = ExcelUtil.getCellStringValue(sheet, i, 4);
                personMessage.setBrithDay(date);

                String perDescription = ExcelUtil.getCellStringValue(sheet, i, 5);
                personMessage.setPerDescription(perDescription);

                list.add(personMessage);

            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return list;

    }

}
