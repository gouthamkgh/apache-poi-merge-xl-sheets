import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class Util {
    
    public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet){
        copySheets(newSheet, sheet, true);
    }
    
    public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet, boolean copyStyle){
        int maxColumnNum = 0;
        Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            XSSFRow srcRow = sheet.getRow(i);
            XSSFRow destRow = newSheet.createRow(i);
            if (srcRow != null) {
                Util.copyRow(sheet, newSheet, srcRow, destRow, styleMap);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    public static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {
        Set<CellRangeAddress> mergedRegions = new TreeSet<CellRangeAddress>();
        destRow.setHeight(srcRow.getHeight());
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            XSSFCell oldCell = srcRow.getCell(j);
            XSSFCell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                copyCell(oldCell, newCell, styleMap);
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(), (short)oldCell.getColumnIndex());
                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastRow(), mergedRegion.getLastColumn());
                    if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
                        mergedRegions.add(newMergedRegion);
                        destSheet.addMergedRegion(newMergedRegion);
                    }
                }
            }
        }
        
    }
    
    public static void copyCell(XSSFCell oldCell, XSSFCell newCell, Map<Integer, XSSFCellStyle> styleMap) {
        if(styleMap != null) {
            if(oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()){
                newCell.setCellStyle(oldCell.getCellStyle());
            } else{
                int stHashCode = oldCell.getCellStyle().hashCode();
                XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
                if(newCellStyle == null){
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                    styleMap.put(stHashCode, newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
        switch(oldCell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }
        
    }
    
    public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    private static boolean isNewMergedRegion(CellRangeAddress newMergedRegion, Collection<CellRangeAddress> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }
    
}