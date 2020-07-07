package usingArray;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

/*
 * Copyright 2020 O3Wallet Technology PTY LTD.
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
public class dataDriven {
    public static void main(String[] args) {

    }

        /**
         * Identify Columns of testcase by scanning the First row
         *
         * Once column is idnetifed then scan entire testcase column to identify PURCHASE testcase row
         *
         * After you grab purchase testcase row pull all the data of row feed into test
         * @return
         * @param add_profile
         */

       public ArrayList<String> getData(String add_profile) throws IOException {
        ArrayList<String> a = new ArrayList<String>();
        // Get excel file using FileInputStream
        FileInputStream file = new FileInputStream("/Users/apple/Documents/TestData.xlsx");

        // Configure XSSFWorkbook object to access specific workbook
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // Read sheet1(Student) of the workbook from all sheets
        int sheets = workbook.getNumberOfSheets();
        for(int i=0; i<sheets; i++){
            if(workbook.getSheetName(i).equalsIgnoreCase("Student")) {
                XSSFSheet sheet = workbook.getSheetAt(i);

                //Identify Columns of testcase by scanning the First row
                Iterator<Row> rows = sheet.iterator();  // Sheet is collection of rows
                Row firstrow = rows.next();

                // scan entire row by read each and every cell
                Iterator<Cell> cell = firstrow.cellIterator();  // Sheet is collection of cells
                int k =0;
                int column = 0;

                while (cell.hasNext()){    // hasNext() only checks that whether it has next cell present or not
                  Cell value = cell.next();   // next() moves control to the next cell
                   if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
                   {
                        // desired Column Number
                       column = k;
                   }
                   k++;
                }
                System.out.println(column);

                //Once column is idnetifed then scan entire testcase column to identify PURCHASE testcase row
                while (rows.hasNext()){
                   Row r = rows.next();
                   if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")){

                       // After you grab purchase testcase row pull all the data of row feed into test
                       Iterator<Cell> cv = r.cellIterator();
                       while(cv.hasNext()){
                          a.add(cv.next().getStringCellValue());
                       }
                   }

                }

            }
        }
            return a;
    }
}
