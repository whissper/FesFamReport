package fesfamreport;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SAV2
 */
public class ReportEngine 
{
    String directoryPlace = "D:/";
    
    String pacientsNamePrefix = "LM110135S";
    String uslugiNamePrefix = "HM110135S";
    
    String namePostfixStart = "11001_";
    String namePostfixEnd = "151001";
    
    String pacientsName = pacientsNamePrefix + namePostfixStart + namePostfixEnd;
    String uslugiName = uslugiNamePrefix + namePostfixStart + namePostfixEnd;
        
    int firstPerson = 7;
    int lastPerson = 8;
        
    int firstUsluga = 7;
    int lastUsluga = 8;
    
    //Constructor - start:
    public ReportEngine()
    {
        
    }
    //Constructor - end;
    
    //methods start;
    public void setDirectoryPlace(String val)
    {
        this.directoryPlace = safeTrim(val) + "/";
    }
    
    public void setNamePostfix(String valStart, String valEnd)
    {
        this.namePostfixStart = safeTrim(valStart);
        this.namePostfixEnd = safeTrim(valEnd);
        this.pacientsName = pacientsNamePrefix + namePostfixStart + namePostfixEnd;
        this.uslugiName = uslugiNamePrefix + namePostfixStart + namePostfixEnd;
    }
    
    public void setFirstLastRows(int firstVal, int lastVal)
    {
        this.firstPerson = firstVal;
        this.lastPerson = lastVal;
        
        this.firstUsluga = firstVal;
        this.lastUsluga = lastVal;
    }
    //methods end;
    
    
    //safeTrim
    private String safeTrim(String value)
    {
        if((value!=null) && (value.length()>0))
        {
            return value.trim();
        }
        else
        {
            return value;
        }
    }
    
    //getCellValue
    private String getCellValue(Cell cellObject)
    {
        String cellValue = "";
        
        if(cellObject != null)
        {
            switch (cellObject.getCellType()) {
                case XSSFCell.CELL_TYPE_NUMERIC:  cellValue += (int)cellObject.getNumericCellValue();
                                                  break;
                case XSSFCell.CELL_TYPE_STRING:   cellValue = cellObject.getStringCellValue();
                                                  break;
                }
        }
        else
        {
            cellValue = "";
        }
        
        return this.safeTrim(cellValue);
    }
    
    //createReport
    public String createReport()
    {
        String protocol = "OK! Файлы сформированы.";
        
        InputStream in = null;
        XSSFWorkbook wb = null;
        //HSSFWorkbook wb = null;
        
        try 
        {
            in = new FileInputStream(safeTrim(directoryPlace) + "Pacients.xlsx");
            wb = new XSSFWorkbook(in);
            //wb = new HSSFWorkbook(in);
        }
        catch(IOException e)
        {
            //e.printStackTrace();
            protocol = "An Error occured during creating of FileInputStream or WorkBook on Pacients stage: \n" + e;
        }

        Sheet sheet = wb.getSheetAt(0);
        
        String xmlString =  "<?xml version=\"1.0\" encoding=\"WINDOWS-1251\"?>";
               xmlString += "<PERS_LIST>";
               
        xmlString += "<ZGLV>";
        for(int i=1-1; i<4; i++)
        { 
            Row row = sheet.getRow(i);
            
            String cellvalue = getCellValue(row.getCell(CellReference.convertColStringToIndex("B")));
            
            switch (i) {
            case 0:  xmlString += "<VERSION>" + cellvalue + "</VERSION>";
                     break;
            case 1:  xmlString += "<DATA>" + cellvalue + "</DATA>";
                     break;
            case 2:  xmlString += "<FILENAME>" + cellvalue + "</FILENAME>";
                     break;
            case 3:  xmlString += "<FILENAME1>" + cellvalue + "</FILENAME1>";
                     break;
            }
            
        }
        xmlString += "</ZGLV>";
        
        for(int i=firstPerson-1; i<lastPerson; i++)
        {
            Row row = sheet.getRow(i);
            
            xmlString += "<PERS xmlns:rk=\"http://komifoms.ru\">";
            
            xmlString += "<ID_PAC>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("B"))) + "</ID_PAC>";
            xmlString += "<FAM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("C"))) + "</FAM>";
            xmlString += "<IM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("D"))) + "</IM>";
            xmlString += "<OT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("E"))) + "</OT>";
            xmlString += "<W>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("F"))) + "</W>";
            xmlString += "<DR>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("G"))) + "</DR>";
            xmlString += "<MR>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("H"))) + "</MR>";
            xmlString += "<DOCTYPE>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("I"))) + "</DOCTYPE>";
            xmlString += "<DOCSER>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("J"))) + "</DOCSER>";
            xmlString += "<DOCNUM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("K"))) + "</DOCNUM>";
            xmlString += "<SNILS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("L"))) + "</SNILS>";
            xmlString += "<OKATOG>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("M"))) + "</OKATOG>";
            xmlString += "<OKATOP>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("N"))) + "</OKATOP>";
            xmlString += "<rk:ADDRES_G>";
                xmlString += "<rk:BOMG>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("O"))) + "</rk:BOMG>";
                xmlString += "<rk:SUBJ>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("P"))) + "</rk:SUBJ>";
                xmlString += "<rk:NPNAME>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Q"))) + "</rk:NPNAME>";
                xmlString += "<rk:UL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("R"))) + "</rk:UL>";
                xmlString += "<rk:DOM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("S"))) + "</rk:DOM>";
                xmlString += "<rk:KV>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("T"))) + "</rk:KV>";
            xmlString += "</rk:ADDRES_G>";
            xmlString += "<rk:ADDRES_P>";
                xmlString += "<rk:SUBJ>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("U"))) + "</rk:SUBJ>";
                xmlString += "<rk:NPNAME>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("V"))) + "</rk:NPNAME>";
                xmlString += "<rk:UL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("W"))) + "</rk:UL>";
                xmlString += "<rk:DOM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("X"))) + "</rk:DOM>";
                xmlString += "<rk:KV>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Y"))) + "</rk:KV>";
            xmlString += "</rk:ADDRES_P>";
            xmlString += "<rk:UNEMP>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Z"))) + "</rk:UNEMP>";
            
            xmlString += "</PERS>";
        }
        
        xmlString += "</PERS_LIST>";
        
        //System.out.println(xmlString);
        try{
            PrintWriter writer = new PrintWriter(safeTrim(directoryPlace)+pacientsName+".xml", "cp1251");
            writer.print(xmlString);
            writer.close();
        }
        catch(Exception e)
        {
            //e.printStackTrace();
            protocol = "An Error occured during saving of XML-file on Pacients stage: \n" + e;
        }
        
        //2-2-2-2-2
        try 
        {
            in = new FileInputStream(safeTrim(directoryPlace) + "Uslugi.xlsx");
            wb = new XSSFWorkbook(in);
            //wb = new HSSFWorkbook(in);
        }
        catch(IOException e)
        {
            //e.printStackTrace();
            protocol = "An Error occured during creating of FileInputStream or WorkBook on Uslugi stage: \n" + e;
        }

        sheet = wb.getSheetAt(0);
        
        xmlString  = "<?xml version=\"1.0\" encoding=\"WINDOWS-1251\"?>";
        xmlString += "<ZL_LIST>";
        
        xmlString += "<ZGLV>";
        for(int i=1-1; i<3; i++)
        { 
            Row row = sheet.getRow(i);
            
            String cellvalue = getCellValue(row.getCell(CellReference.convertColStringToIndex("B")));
            
            switch (i) {
            case 0:  xmlString += "<VERSION>" + cellvalue + "</VERSION>";
                     break;
            case 1:  xmlString += "<DATA>" + cellvalue + "</DATA>";
                     break;
            case 2:  xmlString += "<FILENAME>" + cellvalue + "</FILENAME>";
                     break;
            }
            
        }
        xmlString += "</ZGLV>";
        
        xmlString += "<SCHET xmlns:rk=\"http://komifoms.ru\">";
        Row rowSchet = sheet.getRow(2);
            xmlString += "<CODE>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("D"))) + "</CODE>";
            xmlString += "<CODE_MO>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("E"))) + "</CODE_MO>";
            xmlString += "<YEAR>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("F"))) + "</YEAR>";
            xmlString += "<MONTH>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("G"))) + "</MONTH>";
            xmlString += "<NSCHET>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("H"))) + "</NSCHET>";
            xmlString += "<DSCHET>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("I"))) + "</DSCHET>";
            xmlString += "<rk:PR_NOV>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("J"))) + "</rk:PR_NOV>";
            xmlString += "<PLAT>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("K"))) + "</PLAT>";
            xmlString += "<SUMMAV>" + getCellValue(rowSchet.getCell(CellReference.convertColStringToIndex("L"))) + "</SUMMAV>";    
        xmlString += "</SCHET>";
        
        for(int i=firstUsluga-1; i<lastUsluga; i++)
        {
            Row row = sheet.getRow(i);
            
            xmlString += "<ZAP xmlns:rk=\"http://komifoms.ru\">";
            
            xmlString += "<N_ZAP>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("B"))) + "</N_ZAP>";
            xmlString += "<PR_NOV>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("C"))) + "</PR_NOV>";
            xmlString += "<PACIENT>";
                xmlString += "<ID_PAC>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("D"))) + "</ID_PAC>";
                xmlString += "<rk:ENP>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("E"))) + "</rk:ENP>";
                xmlString += "<VPOLIS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("F"))) + "</VPOLIS>";
                xmlString += "<SPOLIS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("G"))) + "</SPOLIS>";
                xmlString += "<NPOLIS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("H"))) + "</NPOLIS>";
                xmlString += "<SMO>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("I"))) + "</SMO>";
                xmlString += "<NOVOR>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("J"))) + "</NOVOR>";
                xmlString += "<rk:ST_IDENT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("K"))) + "</rk:ST_IDENT>";
                xmlString += "<rk:KEY_IDENT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("L"))) + "</rk:KEY_IDENT>";
            xmlString += "</PACIENT>";
            xmlString += "<SLUCH>";
                xmlString += "<IDCASE>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("M"))) + "</IDCASE>";
                xmlString += "<USL_OK>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("N"))) + "</USL_OK>";
                xmlString += "<VIDPOM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("O"))) + "</VIDPOM>";
                xmlString += "<FOR_POM>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("P"))) + "</FOR_POM>";
                xmlString += "<LPU>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Q"))) + "</LPU>";
                xmlString += "<PODR>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("R"))) + "</PODR>";
                xmlString += "<PROFIL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("S"))) + "</PROFIL>";
                xmlString += "<DET>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("T"))) + "</DET>";
                xmlString += "<NHISTORY>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("U"))) + "</NHISTORY>";
                xmlString += "<DATE_1>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("V"))) + "</DATE_1>";
                xmlString += "<DATE_2>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("W"))) + "</DATE_2>";
                xmlString += "<DS1>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("X"))) + "</DS1>";
                xmlString += "<RSLT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Y"))) + "</RSLT>";
                xmlString += "<ISHOD>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("Z"))) + "</ISHOD>";
                xmlString += "<PRVS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AA"))) + "</PRVS>";
                xmlString += "<IDDOKT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AB"))) + "</IDDOKT>";
                xmlString += "<FIODOKT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AC"))) + "</FIODOKT>";//!
                xmlString += "<IDSP>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AD"))) + "</IDSP>";
                xmlString += "<ED_COL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AE"))) + "</ED_COL>";
                xmlString += "<TARIF>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AF"))) + "</TARIF>";
                xmlString += "<SUMV>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AG"))) + "</SUMV>";
                xmlString += "<FINISH>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AH"))) + "</FINISH>";//!
                xmlString += "<USL>";
                    xmlString += "<IDSERV>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AI"))) + "</IDSERV>";
                    xmlString += "<LPU>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AJ"))) + "</LPU>";
                    xmlString += "<PODR>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AK"))) + "</PODR>";
                    xmlString += "<PROFIL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AL"))) + "</PROFIL>";
                    xmlString += "<DET>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AM"))) + "</DET>";
                    xmlString += "<DATE_IN>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AN"))) + "</DATE_IN>";
                    xmlString += "<DATE_OUT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AO"))) + "</DATE_OUT>";
                    //xmlString += "<DAYSFACT>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AN"))) + "</DAYSFACT>";//---
                    //xmlString += "<FINISH>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AO"))) + "</FINISH>";
                    xmlString += "<PKF>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AP"))) + "</PKF>";
                    xmlString += "<CODE_PK>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AQ"))) + "</CODE_PK>";
                    xmlString += "<DS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AR"))) + "</DS>";
                    xmlString += "<O_KSG>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AS"))) + "</O_KSG>";
                    xmlString += "<KSG>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AT"))) + "</KSG>";
                    xmlString += "<CODE_USL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AU"))) + "</CODE_USL>";
                    xmlString += "<KOL_USL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AV"))) + "</KOL_USL>";
                    xmlString += "<TARIF>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AW"))) + "</TARIF>";
                    xmlString += "<SUMV_USL>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AX"))) + "</SUMV_USL>";
                    xmlString += "<PRVS>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AY"))) + "</PRVS>";
                    xmlString += "<CODE_MD>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("AZ"))) + "</CODE_MD>";
                    xmlString += "<FIO_MD>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("BA"))) + "</FIO_MD>";//!
                    xmlString += "<rk:CODE_SPEC>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("BB"))) + "</rk:CODE_SPEC>";
                    xmlString += "<rk:SELOO>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("BC"))) + "</rk:SELOO>";
                    xmlString += "<rk:IDSP1>" + getCellValue(row.getCell(CellReference.convertColStringToIndex("BD"))) + "</rk:IDSP1>";
                xmlString += "</USL>";
            xmlString += "</SLUCH>";
            
            
            xmlString += "</ZAP>";
        }
        
        xmlString += "</ZL_LIST>";
        
        //System.out.println(xmlString);
        try{
            PrintWriter writer = new PrintWriter(safeTrim(directoryPlace)+uslugiName+".xml", "cp1251");
            writer.print(xmlString);
            writer.close();
        }
        catch(Exception e)
        {
            //e.printStackTrace();
            protocol = "An Error occured during saving of XML-file on Uslugi stage: \n" + e;
        }
        
        return protocol;
    }
    
}
