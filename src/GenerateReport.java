import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class GenerateReport {


    public static void main(String args[]){
        try{

            // Reads input file
            String inputFile = "status.xlsx";
            FileInputStream iINP = new FileInputStream(inputFile);
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(iINP);
            XSSFSheet sheet = inputWorkbook.getSheetAt(0);
            int numRows = sheet.getLastRowNum();

            //reads output file
            String outputFile="Report.xlsx";
            FileInputStream oINP = new FileInputStream(outputFile);
            XSSFWorkbook outputWorkbook = new XSSFWorkbook(oINP);
            int numSheets = outputWorkbook.getNumberOfSheets();
            String weekName = "Week" + numSheets;
            outputWorkbook.createSheet(weekName);

            //writes new week sheet
            FileOutputStream oFOS = new FileOutputStream(outputFile);
            outputWorkbook.write(oFOS);
            System.out.println(weekName + " : sheet created");
            oFOS.close();


            //appends data from input file to week sheet
            XSSFSheet weekSheet = outputWorkbook.getSheet(weekName);
            for(int i =0; i<= numRows; i++){

                Row readRow = sheet.getRow(i);
                Row writeRow = weekSheet.createRow(i);
                int numCols = readRow.getLastCellNum();

                for(int j=0; j<numCols; j++){
                    Cell readCell = readRow.getCell(j);
                    Cell writeCell = writeRow.createCell(j);

                    if(readCell.getCellType() == CellType.NUMERIC){
                        int value = (int) readCell.getNumericCellValue();
                        //System.out.print(value + "--");
                        writeCell.setCellValue(value);
                        weekSheet.autoSizeColumn(j);
                    }else if(readCell.getCellType() == CellType.STRING){
                        //System.out.print(readCell.getStringCellValue() + "--");
                        writeCell.setCellValue(readCell.getStringCellValue());
                        weekSheet.autoSizeColumn(j);
                    }else if(readCell.getCellType() == CellType.BOOLEAN){
                        // System.out.print(readCell.getBooleanCellValue() + "--");
                        writeCell.setCellValue(readCell.getBooleanCellValue());
                        weekSheet.autoSizeColumn(j);
                    }
                }
              //  System.out.println();
            }
            //writes data into week sheet
            oFOS = new FileOutputStream(outputFile);
            outputWorkbook.write(oFOS);
            System.out.println(weekName + " : data appended");
            oFOS.close();

            //reads report sheet
            XSSFSheet reportSheet = outputWorkbook.getSheet("Report");
            //calculates stats from week sheet
            for(int row = 0; row <13; row++ ){
                Row readRow = reportSheet.getRow(row);
                Cell writeCell = readRow.createCell(numSheets);

                if(row == 0){
                    writeCell.setCellValue(weekName);
                }else if(row == 1){
                    String formula = "COUNTIFS("+weekName+"!D2:D"+(numRows+1) +",\"SCHEDULED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 2){
                    String formula = "COUNTIFS("+weekName+"!D2:D"+(numRows+1) +",\"COLLECTED\","+weekName+"!E2:E"+(numRows+1)+",\"\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 3){
                    String formula = "COUNTIFS("+weekName+"!D2:D"+(numRows+1) +",\"CORRUPTED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 4){
                    String formula = "COUNTIFS("+weekName+"!E2:E"+(numRows+1) +",\"RAW\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 5){
                    String formula = "COUNTIFS("+weekName+"!E2:E"+(numRows+1) +",\"ASSIGNED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 6){
                    String formula = "COUNTIFS("+weekName+"!E2:E"+(numRows+1) +",\"LABELED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 7){
                    String formula = "COUNTIFS("+weekName+"!E2:E"+(numRows+1) +",\"RE-LABEL\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 8){
                    String formula = "COUNTIFS("+weekName+"!E2:E"+(numRows+1) +",\"GOOD\","+weekName+"!F2:F"+(numRows+1)+",\"\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 9){
                    String formula = "COUNTIFS("+weekName+"!F2:F"+(numRows+1) +",\"QUEUED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 10){
                    String formula = "COUNTIFS("+weekName+"!F2:F"+(numRows+1) +",\"UPLOADED\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else if(row == 11){
                    String formula = "COUNTIFS("+weekName+"!D2:D"+(numRows+1) +",\"\")";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }else{
                    String formula = "COUNT("+weekName+"!A2:A3000)";
                    writeCell.setCellType(CellType.FORMULA);
                    writeCell.setCellFormula(formula);
                }
            }
            //writes stats into report sheet
            oFOS = new FileOutputStream(outputFile);
            outputWorkbook.write(oFOS);
            System.out.println(weekName + " : Report appended");
            oFOS.close();

            /*****************************************************************************************************/

            final int NUM_OF_COL = reportSheet.getRow(0).getLastCellNum();
            final int NUM_OF_ROW = 12;

            Drawing drawing = reportSheet.createDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 20, 20, 40);

            Chart chart = ((XSSFDrawing) drawing).createChart(anchor);
            ChartLegend legend = chart.getOrCreateLegend();
            legend.setPosition(LegendPosition.RIGHT);
            LineChartData data = chart.getChartDataFactory().createLineChartData();
            ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
            ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            ChartDataSource<Number> xs = DataSources.fromNumericCellRange(reportSheet, new CellRangeAddress(0, NUM_OF_ROW - 1, 0, 0));

            for(int i=1; i<NUM_OF_COL; i++){
                ChartDataSource<Number> ys = DataSources.fromNumericCellRange(reportSheet, new CellRangeAddress(0, NUM_OF_ROW - 1, i, i));
                LineChartSeries series1 = data.addSeries(xs, ys);
                series1.setTitle(weekName);
            }

            chart.plot(data, bottomAxis, leftAxis);

            XSSFChart xssfChart = (XSSFChart) chart;
            CTPlotArea plotArea = xssfChart.getCTChart().getPlotArea();
            plotArea.getLineChartArray()[0].getSmooth();
            CTBoolean ctBool = CTBoolean.Factory.newInstance();
            ctBool.setVal(false);
            plotArea.getLineChartArray()[0].setSmooth(ctBool);
            for (CTLineSer ser : plotArea.getLineChartArray()[0].getSerArray()) {
                ser.setSmooth(ctBool);
            }

            oFOS = new FileOutputStream(outputFile);
            outputWorkbook.write(oFOS);
            System.out.println(weekName + " : Graph appended");
            oFOS.close();


            System.out.println("Press any key and press enter to quit");
            Scanner in = new Scanner(System.in);
            String exit = in.next();
            System.exit(1);

        }catch(IOException ex){
            ex.printStackTrace();
            System.out.println("Report the above error to admin");
            System.out.println("Press any key and press enter to quit");
            Scanner in = new Scanner(System.in);
            String exit = in.next();
            System.exit(1);
        }
    }
}
