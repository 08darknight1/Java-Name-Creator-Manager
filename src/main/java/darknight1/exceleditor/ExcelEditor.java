package darknight1.exceleditor;

import java.io.*;
import java.util.*;

import java.nio.file.*;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelEditor {

    public static void main(String[] args) throws FileNotFoundException, IOException
    {
        boolean running = true;
        
        var currentProjectPath = System.getProperty("user.dir") + "\\" + "ExcelFiles";
        
        File ExcelDirectory = new File(currentProjectPath);
        
        CheckForFolderCreation(ExcelDirectory);
        
        CheckAndCreateExcelFile(currentProjectPath);
            
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(currentProjectPath + "\\ExcelNames.xlsx"));
        HSSFSheet sheet = workbook.getSheetAt(0);
            
        while(running)
        {
            System.out.println("\nWhat do you want to do?");
            System.out.println("0 - Add new Name to file!");
            System.out.println("1 - Add new NickName to file!");
            System.out.println("2 - Add new SurName to file!");
            System.out.println("3 - Export file to somewhere else as CSV!");
            System.out.println("4 - Exit program!");
            System.out.print("//-:");
            
            Scanner scanner = new Scanner(System.in);
            
            int menuOption = scanner.nextInt();

            if(menuOption <= 2)
            {
                while(true)
                {
                    workbook = new HSSFWorkbook(new FileInputStream(currentProjectPath + "\\ExcelNames.xlsx"));
                    
                    sheet = workbook.getSheetAt(0);

                    List<String> namesAlreadyChosen = new ArrayList<>();

                    boolean addingNamesToList = true;

                    int index = 1;

                    while(addingNamesToList)
                    {                
                        if(sheet.getRow(index) == null)
                        {
                            sheet.createRow(index);
                            //System.out.println("\nNumber of names already found: " + namesAlreadyChosen.size());
                            break;
                        }
                        else
                        {
                            if(sheet.getRow(index).getCell(menuOption) != null)
                            {
                                if(sheet.getRow(index).getCell(menuOption).toString() != "")
                                {
                                    namesAlreadyChosen.add(sheet.getRow(index).getCell(menuOption).toString());
                                    //System.out.println("\nAdding new name to current number of names list!");
                                }
                            }
                            else
                            {
                                break;
                            }
                        }

                        index++;
                    }

                    System.out.println("\nType the new name to add or EXIT in full caps to go back to the first menu: ");
                    System.out.print("//-:");

                    scanner = new Scanner(System.in);

                    String newName = scanner.nextLine();

                    boolean registerNewName = true;
                    
                    if(newName.equals("EXIT"))
                    {
                        break;
                    }

                    if(namesAlreadyChosen.size() > 0){
                        for(int x = 0; x < namesAlreadyChosen.size(); x++){
                            //System.out.println("Name chosen: " + newName + " | Name to compare: " + namesAlreadyChosen.get(x));
                            if(newName.equals(namesAlreadyChosen.get(x))){
                                System.out.println("\nName already on the list!");
                                registerNewName = false;
                                break;
                            }
                        }
                    }

                    if(registerNewName){
                        FileOutputStream fileOut = new FileOutputStream(currentProjectPath + "\\ExcelNames.xlsx");

                        System.out.println("Current index value: " + index);

                        sheet.getRow(index).createCell(menuOption);
                        sheet.getRow(index).getCell(menuOption).setCellValue(newName);

                        sheet.autoSizeColumn(0);
                        sheet.autoSizeColumn(1);                       
                        sheet.autoSizeColumn(2);
                        
                        workbook.write(fileOut);
                        fileOut.close();
                        workbook.close();
                    }
                }
            }
            else if(menuOption == 3)
            {
                System.out.println("\nType the new path to send the file to: ");
                System.out.print("//-:");

                scanner = new Scanner(System.in);

                String destination = scanner.nextLine();
                
                FileWriter out = new FileWriter(currentProjectPath + "\\ExcelNames.csv");
                
                CSVPrinter csvPrinter = new CSVPrinter(out, CSVFormat.DEFAULT.withDelimiter(';'));
                
                workbook = new HSSFWorkbook(new FileInputStream(currentProjectPath + "\\ExcelNames.xlsx"));
                sheet = workbook.getSheetAt(0);
                
                if(workbook != null)
                {
                    Iterator<Row> rowIterator = sheet.rowIterator();
                    
                    while (rowIterator.hasNext())
                    {
                        Row row1 = rowIterator.next();
                        Iterator<Cell> cellIterator = row1.cellIterator();
                        
                        while (cellIterator.hasNext())
                        {
                            Cell cell = cellIterator.next();
                            csvPrinter.print(cell.getStringCellValue());
                        }
                        
                        csvPrinter.println();
                    }
                }
                
                csvPrinter.flush();
                csvPrinter.close();
                
                Path destinationPath = Paths.get(destination + "\\ExcelNames.csv");
                
                Path source = Paths.get(currentProjectPath + "\\ExcelNames.csv");
                
                System.out.println("Copiando file de: " + source.toString()+ " para: " + destinationPath.toString());
                
                Files.copy(source, destinationPath, StandardCopyOption.REPLACE_EXISTING);
                
                Files.delete(source);
            }
            else
            {
                running = false;     
            }
        }
    }
    
    private static void CheckForFolderCreation(File currentFile)
    {
        if(!currentFile.isDirectory()){
            try{
                if(currentFile.mkdir())
                { 
                    System.out.println("New folder created in this projects path!");
                }
                else
                {
                    System.out.println("Folder wasnt able to be created!");
                }
            } catch(Exception e) {
                e.printStackTrace();
            }
        }
        else
        {
            System.out.println("Folder already exists!");
        }
    }
    
    private static void CheckAndCreateExcelFile(String path)
    {
        File test = new File(path + "\\ExcelNames.xlsx");
        
        if(!test.isFile()){
            try
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFSheet sheet = workbook.createSheet("FirstSheet");
                HSSFRow row = sheet.createRow(0);
            
                row.createCell(0).setCellValue("Names");
                row.createCell(1).setCellValue("NickNames");
                row.createCell(2).setCellValue("SurNames");
                
                Font font = workbook.createFont();
                
                font.setBold(true);
                
                CellStyle cellStyle = workbook.createCellStyle();
                
                cellStyle.setFont(font);
                
                for(int x = 0; x < 3; x++){
                    row.getCell(x).setCellStyle(cellStyle);
                }
            
                FileOutputStream fileOut = new FileOutputStream(path + "\\ExcelNames.xlsx");
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
            
                System.out.println("Your excel name file has been generated succesfuly!");
            
            } catch ( Exception ex ) {
                System.out.println(ex);
            }
        }
        else
        {
            System.out.println("File already exists! No need to create a new one!");
        }
    }
}
